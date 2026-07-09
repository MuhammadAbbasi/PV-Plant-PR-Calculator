import os
import sqlite3
import json
import re
import traceback
from datetime import datetime
from .parser import parse_mother_file, parse_daily_file

def get_db_connection(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def init_db(db_path):
    """Initializes the database schema if it doesn't exist."""
    conn = get_db_connection(db_path)
    cursor = conn.cursor()
    
    # 1. File metadata table for cache tracking
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS file_meta (
        filepath TEXT PRIMARY KEY,
        file_type TEXT, -- 'mother' or 'daily'
        last_modified REAL,
        status TEXT,
        error_msg TEXT
    )
    """)
    
    # 2. Monthly summaries (aggregated daily data from Mother files)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS monthly_summaries (
        date TEXT PRIMARY KEY, -- YYYY-MM-DD
        irradiance_tx1 REAL,
        irradiance_tx3 REAL,
        irradiance_ref REAL,
        energy REAL,
        pr_total REAL,
        pr_scada REAL,
        pr_vcom REAL,
        pr_compensated REAL,
        availability REAL,
        loss_tx1 REAL,
        loss_tx2 REAL,
        loss_tx3 REAL,
        inverter_prs TEXT -- JSON string
    )
    """)
    
    # 3. Daily parameters/summaries from daily files (PVSyst targets, raw PR, compensated PR, tolerances)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS daily_summaries (
        date TEXT PRIMARY KEY, -- YYYY-MM-DD
        total_values REAL,
        valid_poa_values REAL,
        pvsyst_pr_target REAL,
        raw_pr REAL,
        average_pr REAL,
        uncompensated_pr REAL,
        irr_tolerance REAL,
        min_irr_threshold REAL,
        compensated_pr REAL
    )
    """)
    
    # 4. 15-minute intervals from daily child files
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS daily_intervals (
        datetime_str TEXT PRIMARY KEY, -- YYYY-MM-DD HH:MM:SS
        date TEXT,
        time TEXT, -- HH:MM:SS
        irr_tx1_w REAL,
        irr_tx1_kwh REAL,
        irr_tx3_w REAL,
        irr_tx3_kwh REAL,
        irr_ref REAL,
        energy REAL,
        active_power_regulation REAL,
        loss_tx1 REAL,
        loss_tx2 REAL,
        loss_tx3 REAL,
        inverter_powers TEXT, -- JSON string
        inverter_statuses TEXT -- JSON string
    )
    """)
    
    # Create indexes for fast querying
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_intervals_date ON daily_intervals(date)")
    
    conn.commit()
    conn.close()

def get_sync_status(db_path):
    """Returns the current status of file synchronization."""
    conn = get_db_connection(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM file_meta")
    total_files = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM file_meta WHERE status = 'synced'")
    synced_files = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM file_meta WHERE status = 'error'")
    error_files = cursor.fetchone()[0]
    
    cursor.execute("SELECT filepath, error_msg FROM file_meta WHERE status = 'error'")
    errors = [{"file": row[0], "error": row[1]} for row in cursor.fetchall()]
    
    conn.close()
    
    return {
        "total_files": total_files,
        "synced_files": synced_files,
        "error_files": error_files,
        "errors": errors
    }

def sync_reports(reports_dir, db_path, progress_callback=None):
    """
    Scans reports_dir and incrementally parses files that have changed.
    """
    if not os.path.exists(reports_dir):
        print(f"Directory not found: {reports_dir}")
        return
        
    init_db(db_path)
    
    # Find all monthly directories (YYYY MM)
    month_pattern = re.compile(r"^\d{4}\s\d{2}$")
    subdirs = os.listdir(reports_dir)
    month_dirs = [d for d in subdirs if month_pattern.match(d) and os.path.isdir(os.path.join(reports_dir, d))]
    month_dirs.sort()
    
    files_to_process = []
    
    # Scan files first to check mtime
    conn = get_db_connection(db_path)
    cursor = conn.cursor()
    
    # Load existing meta cache
    cursor.execute("SELECT filepath, last_modified FROM file_meta")
    meta_cache = {row[0]: row[1] for row in cursor.fetchall()}
    
    for md in month_dirs:
        year_str, month_str = md.split()
        md_path = os.path.join(reports_dir, md)
        calc_dir = os.path.join(md_path, "PR CALCOLO FILE")
        
        if not os.path.exists(calc_dir):
            continue
            
        for f in os.listdir(calc_dir):
            if not f.endswith(".xlsx") or f.startswith("~$"):
                continue
                
            f_path = os.path.abspath(os.path.join(calc_dir, f))
            mtime = os.path.getmtime(f_path)
            
            is_mother = f.startswith("00 PR_recalculation_")
            is_daily = f.startswith("PR_recalculation_") and not is_mother
            
            if not (is_mother or is_daily):
                continue
                
            # If not in cache, or mtime is different, mark for processing
            if f_path not in meta_cache or meta_cache[f_path] != mtime:
                # If it's a daily file, try to extract the day from the filename
                date_str = None
                if is_daily:
                    day_match = re.search(r"PR_recalculation_(\d{1,2})_", f)
                    if day_match:
                        day_val = int(day_match.group(1))
                        date_str = f"{year_str}-{month_str}-{day_val:02d}"
                        
                files_to_process.append({
                    "path": f_path,
                    "type": "mother" if is_mother else "daily",
                    "mtime": mtime,
                    "date_str": date_str
                })
                
    conn.close()
    
    total_to_process = len(files_to_process)
    if total_to_process == 0:
        print("Database is up-to-date. No files to process.")
        return
        
    print(f"Syncing database: {total_to_process} files need parsing...")
    
    # Process files
    conn = get_db_connection(db_path)
    cursor = conn.cursor()
    
    for idx, item in enumerate(files_to_process):
        filepath = item["path"]
        ftype = item["type"]
        mtime = item["mtime"]
        date_str = item["date_str"]
        
        if progress_callback:
            progress_callback(idx + 1, total_to_process, filepath)
            
        try:
            if ftype == "mother":
                print(f"Parsing Mother: {os.path.basename(filepath)}")
                days = parse_mother_file(filepath)
                # Save each day
                for day_data in days:
                    cursor.execute("""
                    INSERT OR REPLACE INTO monthly_summaries (
                        date, irradiance_tx1, irradiance_tx3, irradiance_ref, energy,
                        pr_total, pr_scada, pr_vcom, pr_compensated, availability,
                        loss_tx1, loss_tx2, loss_tx3, inverter_prs
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        day_data["date"], day_data["irradiance_tx1"], day_data["irradiance_tx3"],
                        day_data["irradiance_ref"], day_data["energy"], day_data["pr_total"],
                        day_data["pr_scada"], day_data["pr_vcom"], day_data["pr_compensated"],
                        day_data["availability"], day_data["loss_tx1"], day_data["loss_tx2"],
                        day_data["loss_tx3"], json.dumps(day_data["inverter_prs"])
                    ))
                    
            elif ftype == "daily" and date_str:
                print(f"Parsing Daily ({date_str}): {os.path.basename(filepath)}")
                daily_data = parse_daily_file(filepath, date_str)
                summary = daily_data["summary"]
                intervals = daily_data["intervals"]
                
                # 1. Insert daily summary
                cursor.execute("""
                INSERT OR REPLACE INTO daily_summaries (
                    date, total_values, valid_poa_values, pvsyst_pr_target,
                    raw_pr, average_pr, uncompensated_pr, irr_tolerance,
                    min_irr_threshold, compensated_pr
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    date_str, summary["total_values"], summary["valid_poa_values"],
                    summary["pvsyst_pr_target"], summary["raw_pr"], summary["average_pr"],
                    summary["uncompensated_pr"], summary["irr_tolerance"],
                    summary["min_irr_threshold"], summary["compensated_pr"]
                ))
                
                # 2. Clear old intervals for this date to avoid duplicates
                cursor.execute("DELETE FROM daily_intervals WHERE date = ?", (date_str,))
                
                # 3. Insert new intervals
                for interval in intervals:
                    cursor.execute("""
                    INSERT OR REPLACE INTO daily_intervals (
                        datetime_str, date, time, irr_tx1_w, irr_tx1_kwh,
                        irr_tx3_w, irr_tx3_kwh, irr_ref, energy,
                        active_power_regulation, loss_tx1, loss_tx2, loss_tx3,
                        inverter_powers, inverter_statuses
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        interval["datetime_str"], date_str, interval["time"],
                        interval["irr_tx1_w"], interval["irr_tx1_kwh"],
                        interval["irr_tx3_w"], interval["irr_tx3_kwh"],
                        interval["irr_ref"], interval["energy"],
                        interval["active_power_regulation"],
                        interval["loss_tx1"], interval["loss_tx2"], interval["loss_tx3"],
                        json.dumps(interval["inverter_powers"]),
                        json.dumps(interval["inverter_statuses"])
                    ))
            
            # Record success in file_meta
            cursor.execute("""
            INSERT OR REPLACE INTO file_meta (filepath, file_type, last_modified, status, error_msg)
            VALUES (?, ?, ?, 'synced', NULL)
            """, (filepath, ftype, mtime))
            conn.commit()
            
        except Exception as e:
            traceback.print_exc()
            error_msg = str(e)
            print(f"Error parsing {filepath}: {error_msg}")
            # Record error in file_meta
            cursor.execute("""
            INSERT OR REPLACE INTO file_meta (filepath, file_type, last_modified, status, error_msg)
            VALUES (?, ?, ?, 'error', ?)
            """, (filepath, ftype, mtime, error_msg))
            conn.commit()
            
    conn.close()
    print("Database sync completed.")
