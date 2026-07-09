import os
import sys
import threading
from fastapi import FastAPI, BackgroundTasks, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
import sqlite3
import json

from .database import init_db, sync_reports, get_db_connection, get_sync_status

app = FastAPI(title="PV Mazara 01 PR Dashboard API")

# Enable CORS for local development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Paths configuration
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "pr_dashboard_cache.db")
# Default reports folder is relative to the workspace root: Z:\2025.01 Mazara 01 A2A\03 - REPORT\Report\01 Daily Reports
# Let's support both standard layout and overriding via env var.
DEFAULT_REPORTS_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "..", "..", "01 Daily Reports"))
REPORTS_DIR = os.environ.get("PR_REPORTS_DIR", DEFAULT_REPORTS_DIR)

# Global sync state
sync_progress = {
    "is_syncing": False,
    "current_file": "",
    "processed": 0,
    "total": 0
}

def run_sync_in_background():
    global sync_progress
    sync_progress["is_syncing"] = True
    sync_progress["current_file"] = "Scansione file in corso..."
    sync_progress["processed"] = 0
    sync_progress["total"] = 0
    
    def progress_cb(current, total, filepath):
        sync_progress["processed"] = current
        sync_progress["total"] = total
        sync_progress["current_file"] = os.path.basename(filepath)
        
    try:
        init_db(DB_PATH)
        sync_reports(REPORTS_DIR, DB_PATH, progress_callback=progress_cb)
    except Exception as e:
        print(f"Background sync error: {e}")
    finally:
        sync_progress["is_syncing"] = False
        sync_progress["current_file"] = ""

@app.on_event("startup")
def startup_event():
    # Initialize DB and start sync in a separate thread
    init_db(DB_PATH)
    thread = threading.Thread(target=run_sync_in_background)
    thread.daemon = True
    thread.start()

@app.post("/api/sync")
def trigger_sync(background_tasks: BackgroundTasks):
    if sync_progress["is_syncing"]:
        return {"status": "already_syncing", "progress": sync_progress}
    background_tasks.add_task(run_sync_in_background)
    return {"status": "sync_started"}

@app.get("/api/sync-status")
def get_sync_info():
    db_status = get_sync_status(DB_PATH)
    return {
        "is_syncing": sync_progress["is_syncing"],
        "processed": sync_progress["processed"],
        "total": sync_progress["total"],
        "current_file": sync_progress["current_file"],
        "cache_stats": db_status
    }

@app.get("/api/years")
def get_years():
    try:
        conn = get_db_connection(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT substr(date, 1, 4) as year FROM monthly_summaries ORDER BY year DESC")
        years = [row["year"] for row in cursor.fetchall() if row["year"]]
        conn.close()
        return years
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/months")
def get_months(year: str = Query(..., regex=r"^\d{4}$")):
    try:
        conn = get_db_connection(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT DISTINCT substr(date, 6, 2) as month 
            FROM monthly_summaries 
            WHERE date LIKE ? 
            ORDER BY month ASC
        """, (f"{year}-%",))
        months = [row["month"] for row in cursor.fetchall() if row["month"]]
        conn.close()
        return months
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/monthly-data")
def get_monthly_data(year: str, month: str):
    try:
        target_prefix = f"{year}-{month}-%"
        conn = get_db_connection(DB_PATH)
        cursor = conn.cursor()
        
        # Load daily summaries (from monthly Mother files)
        cursor.execute("""
            SELECT m.*, d.pvsyst_pr_target
            FROM monthly_summaries m
            LEFT JOIN daily_summaries d ON m.date = d.date
            WHERE m.date LIKE ?
            ORDER BY m.date ASC
        """, (target_prefix,))
        
        rows = cursor.fetchall()
        result = []
        for row in rows:
            data = dict(row)
            # Parse inverter PRs from JSON
            if data["inverter_prs"]:
                data["inverter_prs"] = json.loads(data["inverter_prs"])
            else:
                data["inverter_prs"] = {}
            result.append(data)
            
        conn.close()
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/daily-data")
def get_daily_data(date: str):
    try:
        conn = get_db_connection(DB_PATH)
        cursor = conn.cursor()
        
        # 1. Fetch daily parameters
        cursor.execute("SELECT * FROM daily_summaries WHERE date = ?", (date,))
        summary_row = cursor.fetchone()
        
        # If no daily summary, fallback to monthly aggregated row if it exists
        summary = {}
        if summary_row:
            summary = dict(summary_row)
        else:
            cursor.execute("SELECT * FROM monthly_summaries WHERE date = ?", (date,))
            fallback_row = cursor.fetchone()
            if fallback_row:
                fallback = dict(fallback_row)
                summary = {
                    "date": date,
                    "compensated_pr": fallback["pr_compensated"],
                    "uncompensated_pr": fallback["pr_scada"],
                    "raw_pr": fallback["pr_total"],
                    "pvsyst_pr_target": 83.2, # default or fallback
                    "min_irr_threshold": 50.0,
                    "irr_tolerance": 10.0,
                    "total_values": 96.0,
                    "valid_poa_values": 0.0
                }
            else:
                raise HTTPException(status_code=404, detail=f"No data found for date {date}")
                
        # 2. Fetch 15-minute interval records
        cursor.execute("""
            SELECT * FROM daily_intervals 
            WHERE date = ? 
            ORDER BY time ASC
        """, (date,))
        interval_rows = cursor.fetchall()
        
        intervals = []
        for r in interval_rows:
            int_data = dict(r)
            int_data["inverter_powers"] = json.loads(int_data["inverter_powers"]) if int_data["inverter_powers"] else {}
            int_data["inverter_statuses"] = json.loads(int_data["inverter_statuses"]) if int_data["inverter_statuses"] else {}
            intervals.append(int_data)
            
        conn.close()
        return {
            "summary": summary,
            "intervals": intervals
        }
    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/yearly-summary")
def get_yearly_summary(year: str):
    try:
        conn = get_db_connection(DB_PATH)
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT 
                substr(date, 6, 2) as month,
                SUM(energy) as total_energy,
                AVG(pr_compensated) as avg_pr_compensated,
                AVG(pr_scada) as avg_pr_scada,
                AVG(pr_total) as avg_pr_raw,
                SUM(irradiance_ref) as total_irradiance,
                SUM(loss_tx1) as total_loss_tx1,
                SUM(loss_tx2) as total_loss_tx2,
                SUM(loss_tx3) as total_loss_tx3
            FROM monthly_summaries
            WHERE date LIKE ?
            GROUP BY month
            ORDER BY month ASC
        """, (f"{year}-%",))
        
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# Serve React static app
static_path = os.path.join(BASE_DIR, "frontend", "dist")

if os.path.exists(static_path):
    app.mount("/", StaticFiles(directory=static_path, html=True), name="static")
else:
    @app.get("/")
    def read_root():
        return JSONResponse(
            status_code=200,
            content={
                "message": "PR Dashboard API runs successfully. Frontend is not compiled yet.",
                "reports_directory_configured": REPORTS_DIR,
                "reports_directory_exists": os.path.exists(REPORTS_DIR)
            }
        )
