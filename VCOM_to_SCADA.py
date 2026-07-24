import os
import datetime
import pandas as pd
import numpy as np
import glob

def get_previous_day_meter_value(date_str, vcom_folder):
    """
    Finds the last cumulative SATAC meter reading (MWh) from the previous day's SCADA files.
    If the folder/file doesn't exist, defaults to 0.0 MWh.
    """
    def clean_meter_val(val):
        if val is None or pd.isna(val):
            return None
        if isinstance(val, (int, float)):
            return float(val)
        val_str = str(val).strip().replace(',', '.')
        try:
            return float(val_str)
        except Exception:
            return None

    dt = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    prev_dt = dt - datetime.timedelta(days=1)
    
    # Reports directory is at the grandparent of grandparent of vcom_folder
    # e.g., vcom_folder = .../Report/01 Daily Reports/2026 07/01/vcom
    # day_folder = .../01
    # month_folder = .../2026 07
    # reports_root = .../01 Daily Reports
    day_folder = os.path.dirname(vcom_folder)
    month_folder = os.path.dirname(day_folder)
    reports_root = os.path.dirname(month_folder)
    
    prev_month_str = f"{prev_dt.year} {prev_dt.month:02d}"
    prev_day_str = f"{prev_dt.day:02d}"
    
    prev_day_folder = os.path.join(reports_root, prev_month_str, prev_day_str)
    
    satac_patterns = ["SATAC_Meter_15Min.xlsx", "SATAC_Meter*.xlsx", "*SATAC*.xlsx"]
    if os.path.exists(prev_day_folder):
        for pat in satac_patterns:
            matches = glob.glob(os.path.join(prev_day_folder, pat))
            if matches:
                try:
                    df = pd.read_excel(matches[0])
                    df_meter = df[df['Colonna2'].astype(str).str.strip().str.startswith("Energia attiva prod")]
                    if len(df_meter) > 0:
                        val = df_meter['Colonna3'].values[-1]
                        parsed_val = clean_meter_val(val)
                        if parsed_val is not None:
                            print(f"[{date_str}] Found previous day meter value: {parsed_val}")
                            return parsed_val
                except Exception as ex:
                    print(f"[{date_str}] Error reading previous day SCADA file: {ex}")
                    
    # Also fallback to check test month directory if test_month exists
    # e.g. .../01 Daily Reports/2026 07 VCOM TEST
    if "VCOM TEST" in month_folder:
        prev_day_folder_test = os.path.join(month_folder, prev_day_str)
        if os.path.exists(prev_day_folder_test):
            for pat in satac_patterns:
                matches = glob.glob(os.path.join(prev_day_folder_test, pat))
                if matches:
                    try:
                        df = pd.read_excel(matches[0])
                        df_meter = df[df['Colonna2'].astype(str).str.strip().str.startswith("Energia attiva prod")]
                        if len(df_meter) > 0:
                            val = df_meter['Colonna3'].values[-1]
                            parsed_val = clean_meter_val(val)
                            if parsed_val is not None:
                                print(f"[{date_str}] Found previous day test meter value: {parsed_val}")
                                return parsed_val
                    except Exception as ex:
                        print(f"[{date_str}] Error reading previous day test file: {ex}")
    print(f"[{date_str}] No previous day meter value found; defaulting to 0.0")
    return 0.0

def convert_vcom_to_scada(vcom_folder, output_folder, date_str):
    """
    Converts VCOM 5-minute exports into pseudo-SCADA 15-minute Excel files
    for a given date (formatted as YYYY-MM-DD).
    """
    # Parse date
    dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    date_val = dt.date()
    
    # Locate VCOM files
    vcom_ac_file = None
    vcom_prod_file = None
    for f in os.listdir(vcom_folder):
        if "Potenza_AC" in f and f.endswith(".csv"):
            vcom_ac_file = os.path.join(vcom_folder, f)
        elif "Produzione_energetica" in f and f.endswith(".csv"):
            vcom_prod_file = os.path.join(vcom_folder, f)
            
    if not vcom_ac_file or not vcom_prod_file:
        raise FileNotFoundError(f"VCOM AC or Produzione files not found in {vcom_folder}")
        
    print(f"[{date_str}] Reading VCOM files...")
    # Clean function for VCOM numbers (comma to dot)
    def clean_vcom_val(val):
        if pd.isna(val): return 0.0
        if isinstance(val, str): val = val.replace(',', '.')
        try: return float(val)
        except: return 0.0

    # Load VCOM data
    df_vcom_ac = pd.read_csv(vcom_ac_file, sep='\t', encoding='utf-16', skiprows=1)
    df_vcom_prod = pd.read_csv(vcom_prod_file, sep='\t', encoding='utf-16', skiprows=1)
    
    # Clean data columns
    for col in df_vcom_ac.columns:
        if col != 'Data':
            df_vcom_ac[col] = df_vcom_ac[col].apply(clean_vcom_val)
    for col in df_vcom_prod.columns:
        if col != 'Data':
            df_vcom_prod[col] = df_vcom_prod[col].apply(clean_vcom_val)
            
    # Function to parse VCOM Data (HH.MM) into minutes of the day
    def time_to_minutes(time_str):
        if isinstance(time_str, float):
            time_str = f"{time_str:.2f}"
        parts = str(time_str).split('.')
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        return h * 60 + m

    df_vcom_ac['minutes'] = df_vcom_ac['Data'].apply(time_to_minutes)
    df_vcom_prod['minutes'] = df_vcom_prod['Data'].apply(time_to_minutes)
    
    # Map 5-min intervals to 15-min intervals
    # A 5-min interval maps to the 15-min interval if it's <= that interval's end.
    # e.g., 05, 10, 15 map to 15. 00 maps to 00.
    def map_to_15min_interval(m):
        if m == 0: return 0
        return ((m - 1) // 15 + 1) * 15

    df_vcom_ac['interval_15'] = df_vcom_ac['minutes'].apply(map_to_15min_interval)
    df_vcom_prod['interval_15'] = df_vcom_prod['minutes'].apply(map_to_15min_interval)
    
    # Standard 15-minute time strings (from 00:00 to 23:45)
    time_intervals_min = list(range(0, 24 * 60, 15))
    
    print(f"[{date_str}] Grouping and resampling to 15-minute intervals...")
    
    # 1. PROCESS WEATHER FILES
    # VCOM Irradianza POA (sensore) [W/m²]
    poa_col = [c for c in df_vcom_prod.columns if 'POA (sensore)' in c][0]
    
    weather_15 = []
    for m in time_intervals_min:
        t_str = f"{m//60:02d}:{m%60:02d}:00"
        # For 00:00, use 00:00. For others, average the 3 intervals (e.g. 05, 10, 15 for 15)
        if m == 0:
            sub = df_vcom_prod[df_vcom_prod['minutes'] == 0]
        else:
            sub = df_vcom_prod[df_vcom_prod['interval_15'] == m]
            
        poa_val = sub[poa_col].mean() if len(sub) > 0 else 0.0
        weather_15.append({"time": t_str, "POA": poa_val})
        
    # Generate Weather Excel sheets
    for ws_id in ["TS_01", "TS_03"]:
        rows = []
        for w in weather_15:
            # We must output Colonna1 to Colonna6
            rows.append({
                "Colonna1": f"MW(CA,MW(09,Data_Mod_{ws_id}_Weather_Station_POA.W11))",
                "Colonna2": " POA",
                "Colonna3": w["POA"],
                "Colonna4": "W/m²",
                "Colonna5": date_val,
                "Colonna6": f"{w['time']}.000"
            })
        df_out = pd.DataFrame(rows)
        out_file = os.path.join(output_folder, f"{ws_id}_Weather_15Min.xlsx")
        df_out.to_excel(out_file, index=False, sheet_name=f"{ws_id}_Weather_15Min")
        
    # 2. PROCESS INVERTER FILES
    # VCOM has inverter powers in Watts [W]
    inverters_15 = {f"TX{s}-INV-{i}": [] for s in [1, 2, 3] for i in range(1, 13)}
    
    for m in time_intervals_min:
        if m == 0:
            sub_ac = df_vcom_ac[df_vcom_ac['minutes'] == 0]
        else:
            sub_ac = df_vcom_ac[df_vcom_ac['interval_15'] == m]
            
        for s in [1, 2, 3]:
            for i in range(1, 13):
                col_name = [c for c in df_vcom_ac.columns if f"TX{s}-{i:02d}" in c or f"TX{s}-{i}" in c]
                if not col_name:
                    # try alternative names if any
                    col_name = [c for c in df_vcom_ac.columns if f"TX{s}" in c and f"{i:02d}" in c]
                
                power_w = sub_ac[col_name[0]].mean() if len(sub_ac) > 0 and col_name else 0.0
                power_kw = power_w / 1000.0 # Convert W to kW
                inverters_15[f"TX{s}-INV-{i}"].append(power_kw)
                
    # Write TS_01_Inverter, TS_02_Inverter, TS_03_Inverter
    for s in [1, 2, 3]:
        rows = []
        for idx, m in enumerate(time_intervals_min):
            t_str = f"{m//60:02d}:{m%60:02d}:00"
            for i in range(1, 13):
                prefix = "EA,MW(17" if s == 1 else ("EG,MW(18" if s == 2 else "EM,MW(19")
                inv_tag = f"MW({prefix},Data_Mod_TS_0{s}_Inverter_{i:02d}.I01))"
                
                power_kw = inverters_15[f"TX{s}-INV-{i}"][idx]
                rows.append({
                    "Colonna1": inv_tag,
                    "Colonna2": " Potenza attiva",
                    "Colonna3": power_kw,
                    "Colonna4": "kW",
                    "Colonna5": date_val,
                    "Colonna6": f"{t_str}.000"
                })
        df_out = pd.DataFrame(rows)
        out_file = os.path.join(output_folder, f"TS_0{s}_Inverter_15Min.xlsx")
        df_out.to_excel(out_file, index=False, sheet_name=f"TS_0{s}_Inverter_15Min")

    # 3. PROCESS SATAC METER FILE
    # VCOM integrated power "Potenza [kW]"
    meter_col = [c for c in df_vcom_prod.columns if 'Potenza' in c][0]
    
    cumulative_mwh = get_previous_day_meter_value(date_str, vcom_folder)
    rows = []
    
    for idx, m in enumerate(time_intervals_min):
        t_str = f"{m//60:02d}:{m%60:02d}:00"
        if m == 0:
            sub = df_vcom_prod[df_vcom_prod['minutes'] == 0]
        else:
            sub = df_vcom_prod[df_vcom_prod['interval_15'] == m]
            
        power_kw = sub[meter_col].mean() if len(sub) > 0 else 0.0
        # Energy produced in 15 mins (kWh) = power_kw * 0.25
        energy_kwh = power_kw * 0.25
        energy_mwh = energy_kwh / 1000.0
        
        # Accumulate
        if m > 0:
            cumulative_mwh += energy_mwh
            
        rows.append({
            "Colonna1": "MW(AA,MW(01,Data_Mod_POC_Meter.M30))",
            "Colonna2": " Energia attiva prod.",
            "Colonna3": round(cumulative_mwh, 4),
            "Colonna4": "MWh",
            "Colonna5": date_val,
            "Colonna6": f"{t_str}.000"
        })
        
    df_out = pd.DataFrame(rows)
    out_file = os.path.join(output_folder, "SATAC_Meter_15Min.xlsx")
    df_out.to_excel(out_file, index=False, sheet_name="SATAC_Meter_15Min")
    print(f"[{date_str}] Pseudo-SCADA generation complete. Files saved in: {output_folder}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Convert VCOM 5-min exports to pseudo-SCADA 15-min Excel files.")
    parser.add_argument("--vcom", help="Path to VCOM folder containing CSV files", required=True)
    parser.add_argument("--out", help="Path to output folder where Excel files will be saved", required=True)
    parser.add_argument("--date", help="Date in YYYY-MM-DD format", required=True)
    args = parser.parse_args()
    
    os.makedirs(args.out, exist_ok=True)
    convert_vcom_to_scada(args.vcom, args.out, args.date)
