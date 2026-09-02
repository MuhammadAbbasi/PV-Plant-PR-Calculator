import os
import glob
import calendar
import datetime
import openpyxl
from excel_engine.excel_com import get_excel_app, open_workbook_writable
from config.constants import ITALIAN_MONTHS_4CHAR, ITALIAN_MONTHS_ABBR, VCOM_LIGHT_ORANGE_COLOR_INT, VCOM_DIFFERENCE_NOTE

def sync_mother_file(calcolo_folder, year_val, month_val, vcom_days=None, vendor_pr_dict=None):
    """
    Synchronizes Mother Excel report (00 PR_recalculation_AGOS.xlsx):
    - Inserts 'Meter Reading [MWh]' before 'Energy (day)'
    - Links daily recalculated values ($L$110 meter, $M$111 energy, PRs)
    - Fills VCOM-processed days with Light Orange fill & attaches note
    - Maintains Row 35 (Day 31) and Row 36 (Summary Totals)
    """
    month_abbr = ITALIAN_MONTHS_ABBR[month_val]
    month_4char = ITALIAN_MONTHS_4CHAR[month_val]
    expected_filename = f"00 PR_recalculation_{month_4char}.xlsx"
    alt_filename = f"00_PR_recalculation_{month_4char}.xlsx"
    
    mother_path = os.path.join(calcolo_folder, expected_filename)
    if not os.path.exists(mother_path) and os.path.exists(os.path.join(calcolo_folder, alt_filename)):
        mother_path = os.path.join(calcolo_folder, alt_filename)
        
    excel = get_excel_app()
    wb_mother = open_workbook_writable(excel, mother_path)
    ws = wb_mother.Sheets('PR_Calc')
    
    num_days = calendar.monthrange(year_val, month_val)[1]
    target_summary_row = 5 + num_days  # Row 36 for 31 days
    
    # Check for Meter Reading column before Energy (day)
    meter_col_idx = None
    energy_col_idx = None
    for c in range(2, 20):
        val_str = str(ws.Cells(4, c).Value or "").strip().lower()
        if "meter reading" in val_str or "lettura contatore" in val_str:
            meter_col_idx = c
        elif "energy (day)" in val_str or "energia" in val_str:
            if not energy_col_idx:
                energy_col_idx = c
                
    if not meter_col_idx and energy_col_idx:
        print(f"DEBUG: Inserting 'Meter Reading' column at position {energy_col_idx}...")
        ws.Columns(energy_col_idx).Insert()
        ws.Cells(4, energy_col_idx).Value = "Meter Reading\n[MWh]"
        
    # Search and adjust summary row position if needed
    current_summary_row = None
    for r in range(30, 42):
        c1_val = str(ws.Cells(r, 1).Value or "")
        if "-" not in c1_val and "/" not in c1_val:
            for c_chk in range(2, 15):
                f_text = str(ws.Cells(r, c_chk).Formula or "").upper()
                if "AVERAGE" in f_text or "SUM" in f_text:
                    current_summary_row = r
                    break
            if current_summary_row is not None:
                break
                
    if current_summary_row is not None:
        if current_summary_row < target_summary_row:
            rows_to_insert = target_summary_row - current_summary_row
            for _ in range(rows_to_insert):
                ws.Rows(current_summary_row).Insert()
                ws.Rows(current_summary_row - 1).Copy(ws.Rows(current_summary_row))
                excel.CutCopyMode = False
                for c in range(1, 65):
                    ws.Cells(current_summary_row, c).Value = None
        elif current_summary_row > target_summary_row:
            rows_to_delete = current_summary_row - target_summary_row
            ws.Rows(f"{target_summary_row}:{current_summary_row-1}").Delete()

    # Set Day dates (5 to 4 + num_days)
    for r in range(5, 5 + num_days):
        day_num = r - 4
        ws.Cells(r, 1).Value = f"{year_val}-{month_val:02d}-{day_num:02d}"

    # Write summary row formulas (Row target_summary_row)
    ws.Cells(target_summary_row, 2).Formula = f"=SUM(B5:B{target_summary_row-1})"
    ws.Cells(target_summary_row, 3).Formula = f"=SUM(C5:C{target_summary_row-1})"
    ws.Cells(target_summary_row, 4).Formula = f"=SUM(D5:D{target_summary_row-1})"
    ws.Cells(target_summary_row, 5).Formula = f"=MAX(E5:E{target_summary_row-1})"
    ws.Cells(target_summary_row, 6).Formula = f"=SUM(F5:F{target_summary_row-1})"
    ws.Cells(target_summary_row, 7).Formula = f"=AVERAGE(G5:G{target_summary_row-1})"
    ws.Cells(target_summary_row, 8).Formula = f"=AVERAGE(H5:H{target_summary_row-1})"
    ws.Cells(target_summary_row, 9).Formula = f"=AVERAGE(I5:I{target_summary_row-1})"
    ws.Cells(target_summary_row, 10).Formula = f'=SUMIF(J5:J{target_summary_row-1},"<>0")/COUNTIF(J5:J{target_summary_row-1},"<>0")'
    ws.Cells(target_summary_row, 11).Formula = f"=SUM(K5:K{target_summary_row-1})"
    ws.Cells(target_summary_row, 12).Formula = f"=SUM(L5:L{target_summary_row-1})"
    ws.Cells(target_summary_row, 13).Formula = f"=SUM(M5:M{target_summary_row-1})"

    # Format VCOM rows & comments
    effective_vcom = set(vcom_days) if vcom_days else set()
    for day_num in range(1, num_days + 1):
        r = 5 + day_num - 1
        is_vcom = (day_num in effective_vcom)
        row_rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, 64))
        date_cell = ws.Cells(r, 1)
        
        if is_vcom:
            row_rng.Interior.Color = VCOM_LIGHT_ORANGE_COLOR_INT
            try:
                if date_cell.Comment is not None:
                    date_cell.Comment.Delete()
                date_cell.AddComment(VCOM_DIFFERENCE_NOTE)
                date_cell.Comment.Visible = False
            except Exception:
                pass
        else:
            if row_rng.Interior.Color == VCOM_LIGHT_ORANGE_COLOR_INT:
                row_rng.Interior.ColorIndex = -4142

    excel.Calculation = -4105  # xlCalculationAutomatic
    wb_mother.Save()
    wb_mother.Close(SaveChanges=True)
    print(f"Mother file '{os.path.basename(mother_path)}' synchronized successfully!")
