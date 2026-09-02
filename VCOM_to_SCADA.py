import os
import datetime
import pandas as pd
import numpy as np
import glob

def clean_vcom_val(val):
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace(',', '.')
    try:
        return float(val_str)
    except Exception:
        return 0.0

def load_vcom_csv(file_path):
    """
    Loads a VCOM CSV file trying multiple encodings and separators.
    Handles skiprows=1 or standard headers.
    """
    for enc in ['utf-16', 'utf-16-le', 'utf-16-be', 'utf-8-sig', 'utf-8', 'latin-1']:
        for sep in ['\t', ',', ';']:
            try:
                # First try skiprows=1 (VCOM export with Periodo header)
                df = pd.read_csv(file_path, sep=sep, encoding=enc, skiprows=1)
                if len(df.columns) > 1 and any(c.lower() in ['data', 'time', 'ora'] for c in df.columns):
                    return df
            except Exception:
                pass
            try:
                # Then try without skiprows
                df = pd.read_csv(file_path, sep=sep, encoding=enc)
                if len(df.columns) > 1 and any(c.lower() in ['data', 'time', 'ora'] for c in df.columns):
                    return df
            except Exception:
                pass
    raise ValueError(f"Impossibile leggere il file CSV VCOM: {file_path}")

def get_previous_day_meter_value(date_str, vcom_folder, output_folder=None):
    """
    Finds the last cumulative SATAC meter reading (MWh) from the previous day's SCADA/VCOM files.
    If the folder/file doesn't exist, defaults to 0.0 MWh.
    """
    dt = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    prev_dt = dt - datetime.timedelta(days=1)
    prev_day_str = f"{prev_dt.day:02d}"
    prev_month_str = f"{prev_dt.year} {prev_dt.month:02d}"
    
    satac_patterns = ["SATAC_Meter_15Min.xlsx", "SATAC_Meter*.xlsx", "*SATAC*.xlsx"]
    
    search_folders = []
    if output_folder:
        m_dir = os.path.dirname(os.path.abspath(output_folder))
        search_folders.append(os.path.join(m_dir, prev_day_str))
        search_folders.append(os.path.join(os.path.dirname(m_dir), prev_month_str, prev_day_str))
        
    if vcom_folder:
        v_abs = os.path.abspath(vcom_folder)
        m_dir_v = os.path.dirname(v_abs)
        search_folders.append(os.path.join(m_dir_v, prev_day_str))
        search_folders.append(os.path.join(os.path.dirname(m_dir_v), prev_month_str, prev_day_str))
        
    for prev_folder in search_folders:
        if os.path.exists(prev_folder):
            for pat in satac_patterns:
                matches = glob.glob(os.path.join(prev_folder, pat))
                for match in matches:
                    try:
                        df = pd.read_excel(match)
                        if 'Colonna2' in df.columns:
                            df_meter = df[df['Colonna2'].astype(str).str.strip().str.startswith("Energia attiva prod")]
                            if len(df_meter) > 0:
                                val = clean_vcom_val(df_meter['Colonna3'].values[-1])
                                if val is not None and val > 0:
                                    print(f"[{date_str}] Found previous day meter value: {val} MWh (from {os.path.basename(match)})")
                                    return val
                    except Exception:
                        pass
                        
    print(f"[{date_str}] No previous day meter value found; defaulting to 0.0 MWh")
    return 0.0

def convert_vcom_to_scada(vcom_folder, output_folder, date_str):
    """
    Converts VCOM 5-minute exports into pseudo-SCADA 15-minute Excel files
    for a given date (formatted as YYYY-MM-DD).
    Generates all 7 required files:
      1. TS_01_Weather_15Min.xlsx
      2. TS_03_Weather_15Min.xlsx
      3. TS_01_Inverter_15Min.xlsx
      4. TS_02_Inverter_15Min.xlsx
      5. TS_03_Inverter_15Min.xlsx
      6. SATAC_Meter_15Min.xlsx
      7. Regolazione_della_potenza_attiva_YYYY_MM_DD.xlsx
    """
    dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    date_val = dt.date()
    date_replaced = date_str.replace("-", "_")
    os.makedirs(output_folder, exist_ok=True)
    
    # Locate VCOM files flexibly, filtering by target date if multiple dates are present in the folder
    patterns = [
        date_str.replace("-", "_"),
        date_str,
        f"{dt.year}_{dt.month:02d}_{dt.day:02d}",
        f"{dt.year}_{dt.month}_{dt.day}",
        f"_{dt.day:02d}_",
        f"_{dt.day:02d}.",
    ]
    
    all_csvs = [f for f in os.listdir(vcom_folder) if f.lower().endswith(".csv")]
    matched_csvs = [f for f in all_csvs if any(pat in f for pat in patterns)]
    if not matched_csvs:
        # Fallback to all csvs in single-day folders (e.g. '05/vcom/')
        matched_csvs = all_csvs

    vcom_ac_file = None
    vcom_prod_file = None
    vcom_active_power_file = None
    vcom_energy_file = None
    
    for f in matched_csvs:
        f_lower = f.lower()
        p = os.path.join(vcom_folder, f)
        if "potenza_ac" in f_lower or "potenza ac" in f_lower:
            vcom_ac_file = p
        elif "produzione_energetica" in f_lower or "produzione energetica" in f_lower:
            vcom_prod_file = p
        elif "potenza_attiva" in f_lower or "potenza attiva" in f_lower:
            vcom_active_power_file = p
        elif "energia" in f_lower and "generata" in f_lower:
            vcom_energy_file = p
        elif "energia" in f_lower:
            vcom_energy_file = p
            
    # Fallback: if prod file is missing but active power file exists, use active power file as prod file
    if not vcom_prod_file and vcom_active_power_file:
        vcom_prod_file = vcom_active_power_file
    if not vcom_active_power_file and vcom_prod_file:
        vcom_active_power_file = vcom_prod_file
        
    if not vcom_ac_file or not vcom_prod_file:
        raise FileNotFoundError(
            f"File VCOM richiesti mancanti in '{vcom_folder}'. "
            f"Trovati: AC={os.path.basename(vcom_ac_file) if vcom_ac_file else 'NO'}, "
            f"Prod/Attiva={os.path.basename(vcom_prod_file) if vcom_prod_file else 'NO'}"
        )
        
    print(f"[{date_str}] Lettura file VCOM da: {vcom_folder}")
    df_vcom_ac = load_vcom_csv(vcom_ac_file)
    df_vcom_prod = load_vcom_csv(vcom_prod_file)
    
    df_vcom_reg = None
    if vcom_active_power_file:
        try:
            df_vcom_reg = load_vcom_csv(vcom_active_power_file)
        except Exception as e:
            print(f"[{date_str}] Avviso: Impossibile leggere file regolazione VCOM: {e}")
            df_vcom_reg = None

    df_vcom_energy = None
    if vcom_energy_file:
        try:
            df_vcom_energy = load_vcom_csv(vcom_energy_file)
            print(f"[{date_str}] File Energia VCOM caricato: {os.path.basename(vcom_energy_file)}")
        except Exception as e:
            print(f"[{date_str}] Avviso: Impossibile leggere file energia VCOM: {e}")
            df_vcom_energy = None

    # Normalize column names (ensure 'Data' is the timestamp column)
    for df in [df_vcom_ac, df_vcom_prod] + ([df_vcom_reg] if df_vcom_reg is not None else []) + ([df_vcom_energy] if df_vcom_energy is not None else []):
        for c in df.columns:
            if c.strip().lower() in ['data', 'time', 'ora', 'timestamp']:
                df.rename(columns={c: 'Data'}, inplace=True)
                break
                
    # Clean numeric columns
    for col in df_vcom_ac.columns:
        if col != 'Data':
            df_vcom_ac[col] = df_vcom_ac[col].apply(clean_vcom_val)
    for col in df_vcom_prod.columns:
        if col != 'Data':
            df_vcom_prod[col] = df_vcom_prod[col].apply(clean_vcom_val)
    if df_vcom_reg is not None:
        for col in df_vcom_reg.columns:
            if col != 'Data':
                df_vcom_reg[col] = df_vcom_reg[col].apply(clean_vcom_val)
    if df_vcom_energy is not None:
        for col in df_vcom_energy.columns:
            if col != 'Data':
                df_vcom_energy[col] = df_vcom_energy[col].apply(clean_vcom_val)
            
    # Function to parse VCOM Data (HH.MM or HH:MM or HH:MM:SS) into minutes of the day
    def time_to_minutes(time_str):
        if isinstance(time_str, (int, float)):
            time_str = f"{time_str:.2f}"
        s = str(time_str).strip().replace(':', '.')
        parts = s.split('.')
        try:
            h = int(parts[0])
            m = int(parts[1]) if len(parts) > 1 else 0
            return h * 60 + m
        except Exception:
            return 0

    df_vcom_ac['minutes'] = df_vcom_ac['Data'].apply(time_to_minutes)
    df_vcom_prod['minutes'] = df_vcom_prod['Data'].apply(time_to_minutes)
    if df_vcom_reg is not None:
        df_vcom_reg['minutes'] = df_vcom_reg['Data'].apply(time_to_minutes)
    if df_vcom_energy is not None:
        df_vcom_energy['minutes'] = df_vcom_energy['Data'].apply(time_to_minutes)
    
    # Map 5-min intervals to 15-min intervals (0, 15, 30, 45, 60, ..., 1440)
    def map_to_15min_interval(m):
        if m == 0: return 0
        return ((m - 1) // 15 + 1) * 15

    df_vcom_ac['interval_15'] = df_vcom_ac['minutes'].apply(map_to_15min_interval)
    df_vcom_prod['interval_15'] = df_vcom_prod['minutes'].apply(map_to_15min_interval)
    if df_vcom_reg is not None:
        df_vcom_reg['interval_15'] = df_vcom_reg['minutes'].apply(map_to_15min_interval)
    if df_vcom_energy is not None:
        df_vcom_energy['interval_15'] = df_vcom_energy['minutes'].apply(map_to_15min_interval)
    
    # Standard 15-minute intervals (96 intervals from 00:00 to 23:45)
    time_intervals_min = list(range(0, 24 * 60, 15))
    
    print(f"[{date_str}] Raggruppamento e ricampionamento a intervalli di 15 minuti...")
    
    # -------------------------------------------------------------
    # 1. PROCESS WEATHER FILES (TS_01_Weather_15Min, TS_03_Weather_15Min)
    # -------------------------------------------------------------
    poa_cols = [c for c in df_vcom_prod.columns if 'poa (sensore)' in c.lower() or 'poa' in c.lower()]
    poa_col = poa_cols[0] if poa_cols else None
    
    weather_15 = []
    for m in time_intervals_min:
        t_str = f"{m//60:02d}:{m%60:02d}:00"
        if m == 0:
            sub = df_vcom_prod[df_vcom_prod['minutes'] == 0]
        else:
            sub = df_vcom_prod[df_vcom_prod['interval_15'] == m]
            
        poa_val = sub[poa_col].mean() if (poa_col and len(sub) > 0) else 0.0
        weather_15.append({"time": t_str, "POA": max(0.0, float(poa_val))})
        
    for ws_id in ["TS_01", "TS_03"]:
        prefix = "CA,MW(09" if ws_id == "TS_01" else "CM,MW(11"
        rows = []
        for w in weather_15:
            rows.append({
                "Colonna1": f"MW({prefix},Data_Mod_{ws_id}_Weather_Station_POA.W11))",
                "Colonna2": " POA",
                "Colonna3": w["POA"],
                "Colonna4": "W/m²",
                "Colonna5": date_val,
                "Colonna6": f"{w['time']}.000"
            })
        df_out = pd.DataFrame(rows)
        out_file = os.path.join(output_folder, f"{ws_id}_Weather_15Min.xlsx")
        df_out.to_excel(out_file, index=False, sheet_name=f"{ws_id}_Weather_15Min")
        
    # -------------------------------------------------------------
    # 2. PROCESS INVERTER FILES (TS_01, TS_02, TS_03 Inverter_15Min)
    # -------------------------------------------------------------
    inverters_15 = {f"TX{s}-INV-{i}": [] for s in [1, 2, 3] for i in range(1, 13)}
    
    for m in time_intervals_min:
        if m == 0:
            sub_ac = df_vcom_ac[df_vcom_ac['minutes'] == 0]
        else:
            sub_ac = df_vcom_ac[df_vcom_ac['interval_15'] == m]
            
        for s in [1, 2, 3]:
            for i in range(1, 13):
                # Search for matching column name
                # Examples: 'Potenza AC (INV TX1-01) [W]', 'INV TX1-01', 'TX1-01', 'TX1-1'
                matches = [c for c in df_vcom_ac.columns if f"TX{s}-{i:02d}" in c or f"TX{s}-{i}" in c or f"tx{s}-{i:02d}" in c.lower()]
                if not matches:
                    matches = [c for c in df_vcom_ac.columns if f"TX{s}" in c and f"{i:02d}" in c]
                
                power_w = sub_ac[matches[0]].mean() if (matches and len(sub_ac) > 0) else 0.0
                power_kw = max(0.0, float(power_w) / 1000.0) # Convert W to kW
                inverters_15[f"TX{s}-INV-{i}"].append(power_kw)
                
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

    # -------------------------------------------------------------
    # 3. PROCESS SATAC METER FILE (SATAC_Meter_15Min)
    # -------------------------------------------------------------
    cumulative_mwh_base = get_previous_day_meter_value(date_str, vcom_folder, output_folder=output_folder)
    rows = []
    
    # Check if df_vcom_energy provides per-inverter daily cumulative energy (Energia generata al giorno [kWh])
    energy_inv_cols = []
    if df_vcom_energy is not None:
        energy_inv_cols = [c for c in df_vcom_energy.columns if c != 'Data' and c != 'minutes' and c != 'interval_15']
        
    if df_vcom_energy is not None and len(energy_inv_cols) > 0:
        print(f"[{date_str}] Calcolo letture SATAC Meter da file Energia VCOM ({len(energy_inv_cols)} inverters)...")
        df_vcom_energy['cum_kwh_plant'] = df_vcom_energy[energy_inv_cols].sum(axis=1)
        
        last_kwh = 0.0
        for idx, m in enumerate(time_intervals_min):
            t_str = f"{m//60:02d}:{m%60:02d}:00"
            if m == 0:
                sub = df_vcom_energy[df_vcom_energy['minutes'] == 0]
            else:
                sub = df_vcom_energy[df_vcom_energy['interval_15'] == m]
                
            if len(sub) > 0:
                last_kwh = float(sub['cum_kwh_plant'].mean())
                
            cumulative_mwh = cumulative_mwh_base + (last_kwh / 1000.0)
            rows.append({
                "Colonna1": "MW(AA,MW(01,Data_Mod_POC_Meter.M30))",
                "Colonna2": " Energia attiva prod.",
                "Colonna3": round(cumulative_mwh, 4),
                "Colonna4": "MWh",
                "Colonna5": date_val,
                "Colonna6": f"{t_str}.000"
            })
    else:
        p_cols = [c for c in df_vcom_prod.columns if 'potenza [kw]' in c.lower() or 'potenza' in c.lower()]
        meter_col = p_cols[0] if p_cols else None
        
        running_mwh = cumulative_mwh_base
        for idx, m in enumerate(time_intervals_min):
            t_str = f"{m//60:02d}:{m%60:02d}:00"
            if m == 0:
                sub = df_vcom_prod[df_vcom_prod['minutes'] == 0]
            else:
                sub = df_vcom_prod[df_vcom_prod['interval_15'] == m]
                
            power_kw = sub[meter_col].mean() if (meter_col and len(sub) > 0) else 0.0
            energy_kwh = max(0.0, float(power_kw) * 0.25)
            energy_mwh = energy_kwh / 1000.0
            
            if m > 0:
                running_mwh += energy_mwh
                
            rows.append({
                "Colonna1": "MW(AA,MW(01,Data_Mod_POC_Meter.M30))",
                "Colonna2": " Energia attiva prod.",
                "Colonna3": round(running_mwh, 4),
                "Colonna4": "MWh",
                "Colonna5": date_val,
                "Colonna6": f"{t_str}.000"
            })
        
    df_out = pd.DataFrame(rows)
    out_file = os.path.join(output_folder, "SATAC_Meter_15Min.xlsx")
    df_out.to_excel(out_file, index=False, sheet_name="SATAC_Meter_15Min")

    # -------------------------------------------------------------
    # 4. PROCESS ACTIVE POWER REGULATION (Regolazione_della_potenza_attiva)
    # -------------------------------------------------------------
    reg_col = None
    if df_vcom_reg is not None:
        for c in df_vcom_reg.columns:
            c_low = c.lower()
            if 'potenza attiva [%]' in c_low or 'valore nominale' in c_low or 'regolazione' in c_low:
                reg_col = c
                break
                
    reg_rows = []
    for m in time_intervals_min:
        t_str = f"{m//60:02d}:{m%60:02d}:00"
        reg_val = 100.0
        if df_vcom_reg is not None and reg_col:
            sub = df_vcom_reg[df_vcom_reg['minutes'] == 0] if m == 0 else df_vcom_reg[df_vcom_reg['interval_15'] == m]
            if len(sub) > 0:
                v = sub[reg_col].mean()
                if not pd.isna(v) and v > 0:
                    reg_val = float(v)
        reg_rows.append({
            "Data": t_str,
            "Unnamed: 1": f"{t_str}.000",
            "Valore nominale potenza attiva [%]": reg_val
        })
        
    df_reg_out = pd.DataFrame(reg_rows)
    reg_out_file = os.path.join(output_folder, f"Regolazione_della_potenza_attiva_{date_replaced}.xlsx")
    df_reg_out.to_excel(reg_out_file, index=False)
    
    print(f"[{date_str}] Generazione pseudo-SCADA VCOM completata con successo (7/7 file creati in: {output_folder})")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Convert VCOM 5-min exports to pseudo-SCADA 15-min Excel files.")
    parser.add_argument("--vcom", help="Path to VCOM folder containing CSV files", required=True)
    parser.add_argument("--out", help="Path to output folder where Excel files will be saved", required=True)
    parser.add_argument("--date", help="Date in YYYY-MM-DD format", required=True)
    args = parser.parse_args()
    
    os.makedirs(args.out, exist_ok=True)
    convert_vcom_to_scada(args.vcom, args.out, args.date)
