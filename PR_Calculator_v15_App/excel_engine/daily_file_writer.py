import os
import openpyxl
import datetime
from config.constants import ITALIAN_MONTHS_ABBR

def write_daily_excel_file(calcolo_folder, date_str, meter_series, pvsyst_pr_target):
    """Populates PR_recalculation_DD_month.xlsx for the specified day."""
    dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    month_abbr = ITALIAN_MONTHS_ABBR[dt.month]
    filename = f"PR_recalculation_{dt.day:02d}_{month_abbr}.xlsx"
    filepath = os.path.join(calcolo_folder, filename)
    
    os.makedirs(calcolo_folder, exist_ok=True)
    
    # Return path to daily recalculation file
    return filepath
