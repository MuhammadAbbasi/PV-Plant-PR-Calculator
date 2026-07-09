import os
import re
import json
import numpy as np
import pandas as pd
from datetime import datetime

def clean_float(val):
    """Safely convert a cell value to a float, handling commas and nan."""
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    # If it is a string (e.g. from Italian Excel formatting with commas)
    if isinstance(val, str):
        val = val.strip().replace(',', '.')
        try:
            return float(val)
        except ValueError:
            return 0.0
    return 0.0

def parse_mother_file(filepath):
    """
    Parses the monthly Mother file (00 PR_recalculation_*.xlsx).
    Returns a list of dicts representing each day's row.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Mother file not found: {filepath}")
        
    # Read the PR_Calc sheet
    df = pd.read_excel(filepath, sheet_name='PR_Calc', header=None)
    
    # Row 4 (index 3) is headers. Let's make sure it contains the columns we expect.
    headers = [str(x).strip() for x in df.iloc[3]]
    
    # We find rows containing data (Row 5 onwards, up to the summary row).
    # Data rows start at index 4. The summary row has formulas like SUM(B5:B...) or text.
    # Typically, the date is in column 0 (index 0).
    data_rows = []
    for r in range(4, len(df)):
        date_val = df.iloc[r, 0]
        if pd.isna(date_val):
            continue
            
        # If the date_val is a string that starts with "Dati" or "Totale" or similar, skip it.
        # Or if it represents a summary row
        if isinstance(date_val, str) and (date_val.lower().startswith('tot') or date_val.lower().startswith('dati') or date_val.lower().startswith('average') or date_val.lower().startswith('sum')):
            continue
            
        # Parse date
        date_str = ""
        if isinstance(date_val, datetime):
            date_str = date_val.strftime("%Y-%m-%d")
        elif isinstance(date_val, pd.Timestamp):
            date_str = date_val.strftime("%Y-%m-%d")
        else:
            # Try to parse string
            try:
                date_str = pd.to_datetime(date_val).strftime("%Y-%m-%d")
            except:
                # If we cannot parse, it might be the summary row, skip it
                continue
        
        # Check if the date string matches YYYY-MM-DD
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", date_str):
            continue
            
        # Extract columns
        # Col B (index 1) to Col M (index 12)
        irr_tx1 = clean_float(df.iloc[r, 1])
        irr_tx3 = clean_float(df.iloc[r, 2])
        irr_ref = clean_float(df.iloc[r, 3])
        energy = clean_float(df.iloc[r, 4])
        pr_total = clean_float(df.iloc[r, 5])
        pr_scada = clean_float(df.iloc[r, 6])
        pr_vcom = clean_float(df.iloc[r, 7])
        pr_compensated = clean_float(df.iloc[r, 8])
        availability = clean_float(df.iloc[r, 9])
        loss_tx1 = clean_float(df.iloc[r, 10])
        loss_tx2 = clean_float(df.iloc[r, 11])
        loss_tx3 = clean_float(df.iloc[r, 12])
        
        # Extract Inverter PRs from Column N (index 13) to AW (index 48)
        # There are 36 inverters
        inverter_prs = {}
        for col_idx in range(13, 49):
            inv_name = headers[col_idx] if col_idx < len(headers) else f"PR_INV_{col_idx-12}"
            # Clean name a bit
            inv_name = inv_name.replace("PR ", "").replace("\n", " ").strip()
            # The value is a decimal e.g. 0.38007, let's keep it as decimal or convert to % depending on need
            # We store it as decimal float
            inverter_prs[inv_name] = clean_float(df.iloc[r, col_idx])
            
        data_rows.append({
            "date": date_str,
            "irradiance_tx1": irr_tx1,
            "irradiance_tx3": irr_tx3,
            "irradiance_ref": irr_ref,
            "energy": energy,
            "pr_total": pr_total,
            "pr_scada": pr_scada,
            "pr_vcom": pr_vcom,
            "pr_compensated": pr_compensated,
            "availability": availability,
            "loss_tx1": loss_tx1,
            "loss_tx2": loss_tx2,
            "loss_tx3": loss_tx3,
            "inverter_prs": inverter_prs
        })
        
    return data_rows

def parse_daily_file(filepath, date_str):
    """
    Parses a daily child file (PR_recalculation_*.xlsx).
    Returns a dict containing:
      - 'summary': high-level parameters (PVSyst target, RAW PR, Compensated PR, etc.)
      - 'intervals': list of dicts representing 15-minute data rows.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Daily file not found: {filepath}")
        
    xls = pd.ExcelFile(filepath)
    
    # We parse the 'PR_Calc' sheet
    df_calc = pd.read_excel(xls, sheet_name='PR_Calc', header=None)
    
    # 1. Read high-level parameters from column BH (index 59)
    # Row indices are 0-indexed (Row 3 = index 2, Row 11 = index 10)
    summary_params = {
        "date": date_str,
        "total_values": clean_float(df_calc.iloc[2, 59]),
        "valid_poa_values": clean_float(df_calc.iloc[3, 59]),
        "pvsyst_pr_target": clean_float(df_calc.iloc[4, 59]),
        "raw_pr": clean_float(df_calc.iloc[5, 59]),
        "average_pr": clean_float(df_calc.iloc[6, 59]),
        "uncompensated_pr": clean_float(df_calc.iloc[7, 59]),
        "irr_tolerance": clean_float(df_calc.iloc[8, 59]),
        "min_irr_threshold": clean_float(df_calc.iloc[9, 59]),
        "compensated_pr": clean_float(df_calc.iloc[10, 59])
    }
    
    # 2. Read 15-minute interval data from rows 15 to 110 (index 14 to 109)
    # We also load the Inverter_data sheet to get the power values.
    df_inv = None
    if 'Inverter_data' in xls.sheet_names:
        df_inv = pd.read_excel(xls, sheet_name='Inverter_data', header=None)
        
    intervals = []
    
    # Extract inverter names from Row 14 header of Inverter_data if available
    inv_power_headers = []
    if df_inv is not None:
        inv_power_headers = [str(x).strip() for x in df_inv.iloc[13]]
        
    # Headers of PR_Calc for status columns
    calc_headers = [str(x).strip() for x in df_calc.iloc[13]]
    
    # Row 15 to 110 are data
    for idx in range(14, 110):
        # Time slot is in Col B (index 1)
        time_slot = df_calc.iloc[idx, 1]
        if pd.isna(time_slot):
            continue
            
        if isinstance(time_slot, datetime):
            time_str = time_slot.strftime("%H:%M:%S")
        elif isinstance(time_slot, str):
            # Clean string e.g. "00:00:00.000" to "00:00:00"
            time_str = time_slot.split('.')[0].strip()
        else:
            # Fallback
            time_str = str(time_slot).split('.')[0].strip()
            
        # Match HH:MM:SS format
        if not re.match(r"^\d{2}:\d{2}:\d{2}$", time_str):
            continue
            
        irr_tx1_w = clean_float(df_calc.iloc[idx, 2])
        irr_tx1_kwh = clean_float(df_calc.iloc[idx, 3])
        irr_tx3_w = clean_float(df_calc.iloc[idx, 4])
        irr_tx3_kwh = clean_float(df_calc.iloc[idx, 5])
        irr_ref = clean_float(df_calc.iloc[idx, 8])  # Column I (index 8) is Reference POA
        meter_energy = clean_float(df_calc.iloc[idx, 12])  # Column M (index 12) is energy
        active_power_reg = clean_float(df_calc.iloc[idx, 13])  # Column N limit ratio
        
        # Losses for TX1, TX2, TX3
        loss_tx1 = clean_float(df_calc.iloc[idx, 26])  # Col AA (index 26)
        loss_tx2 = clean_float(df_calc.iloc[idx, 39])  # Col AN (index 39)
        loss_tx3 = clean_float(df_calc.iloc[idx, 52])  # Col BA (index 52)
        
        # Inverter power values from Inverter_data sheet
        inverter_powers = {}
        if df_inv is not None:
            # Active power columns are:
            # C to N (2 to 13)
            # R to AC (17 to 28)
            # AG to AR (32 to 43)
            for col_idx in list(range(2, 14)) + list(range(17, 29)) + list(range(32, 44)):
                header_name = inv_power_headers[col_idx] if col_idx < len(inv_power_headers) else f"INV_power_{col_idx}"
                clean_name = header_name.replace("Active power ", "").replace("\n", " ").strip()
                inverter_powers[clean_name] = clean_float(df_inv.iloc[idx, col_idx])

        # Derive Inverter status values from active power and irradiance
        # (O-Z, AB-AM, AO-AZ in PR_Calc contain overridden loss values, not status)
        inverter_statuses = {}
        is_sun = (irr_ref * 4000) > 50.0
        for inv_name, power in inverter_powers.items():
            if power > 1.0:
                inverter_statuses[inv_name] = 1.0
            elif is_sun:
                inverter_statuses[inv_name] = 0.0
            else:
                inverter_statuses[inv_name] = 1.0
                
        intervals.append({
            "time": time_str,
            "datetime_str": f"{date_str} {time_str}",
            "irr_tx1_w": irr_tx1_w,
            "irr_tx1_kwh": irr_tx1_kwh,
            "irr_tx3_w": irr_tx3_w,
            "irr_tx3_kwh": irr_tx3_kwh,
            "irr_ref": irr_ref,
            "energy": meter_energy,
            "active_power_regulation": active_power_reg,
            "loss_tx1": loss_tx1,
            "loss_tx2": loss_tx2,
            "loss_tx3": loss_tx3,
            "inverter_powers": inverter_powers,
            "inverter_statuses": inverter_statuses
        })
        
    return {
        "summary": summary_params,
        "intervals": intervals
    }
