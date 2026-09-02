import os
import glob
import datetime
import pandas as pd
import numpy as np
from utils.data_cleaners import clean_float

def get_previous_day_meter_value(date_str, vcom_folder, output_folder=None):
    """Finds the last cumulative SATAC meter reading (MWh) from the previous day."""
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
                                val = clean_float(df_meter['Colonna3'].values[-1])
                                if val > 0:
                                    print(f"[{date_str}] Found previous day meter value: {val} MWh (from {os.path.basename(match)})")
                                    return val
                    except Exception:
                        pass
    print(f"[{date_str}] No previous day meter value found; defaulting to 0.0 MWh")
    return 0.0

def convert_vcom_to_scada(vcom_folder, output_folder, date_str):
    """Converts VCOM 3-file data into 7 pseudo-SCADA 15-minute Excel files."""
    dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    date_formatted = dt.strftime("%Y_%m_%d")
    
    energy_file = os.path.join(vcom_folder, f"Energia_{date_formatted}.csv")
    power_ac_file = os.path.join(vcom_folder, f"Potenza_AC_{date_formatted}.csv")
    power_active_file = os.path.join(vcom_folder, f"Potenza_attiva_{date_formatted}.csv")
    
    if not os.path.exists(energy_file) or not os.path.exists(power_ac_file) or not os.path.exists(power_active_file):
        raise FileNotFoundError(f"One or more VCOM CSV files missing for date {date_str} in '{vcom_folder}'")

    os.makedirs(output_folder, exist_ok=True)
    df_energy = pd.read_csv(energy_file, sep=';', encoding='utf-8')
    df_ac = pd.read_csv(power_ac_file, sep=';', encoding='utf-8')
    df_active = pd.read_csv(power_active_file, sep=';', encoding='utf-8')
    
    # 15-min resample grid
    time_grid = pd.date_range(f"{date_str} 00:00:00", f"{date_str} 23:45:00", freq="15min")
    
    # Calculate cumulative MWh meter series
    prev_meter_mwh = get_previous_day_meter_value(date_str, vcom_folder, output_folder=output_folder)
    inv_cols = [c for c in df_energy.columns if "Energia generata al giorno" in c]
    cum_kwh_series = df_energy[inv_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1).values
    
    # Write SATAC Meter 15Min file
    meter_df = pd.DataFrame({
        'Colonna1': ['MW(AA,MW(01,Data_Mod_POC_Meter.M01))'] * 96,
        'Colonna2': ['Energia attiva prodotta dal gruppo'] * 96,
        'Colonna3': prev_meter_mwh + (cum_kwh_series / 1000.0),
        'Colonna4': ['MWh'] * 96,
        'Colonna5': time_grid,
        'Colonna6': time_grid.strftime("%H:%M:%S.000")
    })
    meter_df.to_excel(os.path.join(output_folder, "SATAC_Meter_15Min.xlsx"), index=False)
    
    print(f"[{date_str}] VCOM pseudo-SCADA files successfully generated in '{output_folder}'.")
