# -*- coding: utf-8 -*-
"""
===============================================================================
SCADA Humidity Report Generator - Mazara 01 (A2A)
GET SRL

ULTIMO AGGIORNAMENTO / LAST UPDATE: 2026-09-02 15:42
===============================================================================
CHANGELOG / STORICO MODIFICHE:
  * 2026-09-02:
    - Integrata icona e logo applicativo personalizzato 'UMIDITÀ' (PNG/ICO).
    - Aggiunta funzione get_resource_path per compatibilità bundle PyInstaller (.exe).
    - Visualizzazione logo nella barra del titolo e nell'header della GUI Tkinter.
    - Compilazione e creazione dell'eseguibile standalone 'Report_Umidita_SCADA.exe'.
    - Impostata selezione automatica predefinita del mese precedente (es. ad
      inizio settembre 2026 viene preselezionato automaticamente '2026 08').
    - Creazione iniziale generatore report umidità SCADA con supporto GUI Tkinter
      e CLI da riga di comando (--month, --all, --folder).
    - Generazione automatica del foglio 'Rapporto Umidità' e del foglio di calcolo
      dettagliato SCADA '<Mese> <Anno>' (es. 'July 2026').
    - Implementazione formule qualità e soglia di accettabilità del 5% tra stazioni
      meteo TS_01 e TS_03 per Umidità Relativa e Assoluta.
===============================================================================

Source Files:
  - TS_01_Weather_1Day.xlsx
  - TS_03_Weather_1Day.xlsx

Output File:
  - YYYY MM Dati Humidita.xlsx (e.g. '2026 07 Dati Humidita.xlsx')
    Contains:
      1. 'Rapporto Umidità': Executive summary table with daily combined values and monthly average.
      2. '<Month Name> <YYYY>' (e.g. 'July 2026'): Detailed SCADA calculation sheet with individual
         station readings, 5% acceptance rate comparisons, Excel formulas, and Excel Table formatting.
"""

import os
import sys
import re
import calendar
from datetime import datetime, date, timedelta
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

# Default base directory for Mazara Daily Reports
DEFAULT_BASE_DIR = r"\\s01\get\2025.01 Mazara 01 A2A\03 - REPORT\Report\01 Daily Reports"

ENGLISH_MONTHS = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December"
}

ITALIAN_MONTHS = {
    1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile",
    5: "Maggio", 6: "Giugno", 7: "Luglio", 8: "Agosto",
    9: "Settembre", 10: "Ottobre", 11: "Novembre", 12: "Dicembre"
}


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


def get_previous_month_str():
    """Returns 'YYYY MM' for the month immediately prior to current date (e.g. '2026 08' in Sept 2026)."""
    today = datetime.now()
    first_of_current = today.replace(day=1)
    prev_month_last_day = first_of_current - timedelta(days=1)
    return f"{prev_month_last_day.year:04d} {prev_month_last_day.month:02d}"


def normalize_var_name(name):
    """Normalize variable names to identify relative and absolute humidity."""
    if not name:
        return ""
    cleaned = str(name).strip().lower()
    cleaned = cleaned.replace("à", "a").replace("ì", "i")
    if "umidit" in cleaned or "humid" in cleaned:
        if "relat" in cleaned:
            return "relative"
        if "assolut" in cleaned or "absolut" in cleaned:
            return "absolute"
    return ""


def parse_scada_weather_file(file_path):
    """
    Parse a SCADA 1-Day Weather Excel file (TS_01 or TS_03).
    Returns a dictionary: {(date_obj, 'relative'): value, (date_obj, 'absolute'): value}
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File SCADA non trovato: {file_path}")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    data = {}
    for r in range(2, ws.max_row + 1):
        var_cell = ws.cell(r, 2).value
        var_type = normalize_var_name(var_cell)
        if not var_type:
            continue

        val = ws.cell(r, 3).value
        dt_val = ws.cell(r, 5).value

        d_key = None
        if isinstance(dt_val, datetime):
            d_key = dt_val.date()
        elif isinstance(dt_val, date):
            d_key = dt_val
        elif isinstance(dt_val, str):
            try:
                d_key = datetime.strptime(dt_val[:10], "%Y-%m-%d").date()
            except ValueError:
                pass

        if d_key:
            num_val = float(val) if val is not None else 0.0
            data[(d_key, var_type)] = num_val

    wb.close()
    return data


def extract_year_month_from_folder(folder_path):
    """Extract (year, month) from folder name like '2026 07' or '2026-07'."""
    folder_name = os.path.basename(os.path.normpath(folder_path))
    m = re.search(r"(\d{4})[\s\-_]+(\d{1,2})", folder_name)
    if m:
        return int(m.group(1)), int(m.group(2))

    ts1_file = os.path.join(folder_path, "TS_01_Weather_1Day.xlsx")
    if os.path.exists(ts1_file):
        try:
            wb = openpyxl.load_workbook(ts1_file, data_only=True, read_only=True)
            ws = wb.active
            for r in range(2, min(ws.max_row + 1, 30)):
                dt = ws.cell(r, 5).value
                if isinstance(dt, (datetime, date)):
                    wb.close()
                    return dt.year, dt.month
            wb.close()
        except Exception:
            pass

    raise ValueError(f"Impossibile determinare anno e mese dalla cartella: {folder_path}")


def compute_combined_humidity(tx1_val, tx3_val, var_type, acceptance_rate=0.05):
    """
    Compute combined humidity value according to standard Mazara rules:
    - Relative humidity: threshold is acceptance_rate * 100 (5.0%).
      If |TX1 - TX3| > 5%, take MAX(TX1, TX3), else AVERAGE(TX1, TX3).
    - Absolute humidity: threshold is acceptance_rate * TX1 (5% of TX1).
      If |TX3 - TX1| > 0.05 * TX1, take MAX(TX1, TX3), else AVERAGE(TX1, TX3).
    """
    t1 = float(tx1_val or 0.0)
    t3 = float(tx3_val or 0.0)

    if var_type == "relative":
        diff = abs(t1 - t3)
        if diff > (acceptance_rate * 100.0):
            return max(t1, t3)
        return (t1 + t3) / 2.0
    else:  # absolute
        diff = abs(t3 - t1)
        if diff > (acceptance_rate * t1):
            return max(t1, t3)
        return (t1 + t3) / 2.0


def generate_humidity_report(month_folder, output_path=None, acceptance_rate=0.05, satac_path=None):
    """
    Generate the Humidity Excel report for a given month folder.
    """
    month_folder = os.path.abspath(month_folder)
    if not os.path.isdir(month_folder):
        raise NotADirectoryError(f"Cartella non trovata: {month_folder}")

    ts1_path = os.path.join(month_folder, "TS_01_Weather_1Day.xlsx")
    ts3_path = os.path.join(month_folder, "TS_03_Weather_1Day.xlsx")

    if not os.path.exists(ts1_path):
        raise FileNotFoundError(f"File TS_01 mancante: {ts1_path}")
    if not os.path.exists(ts3_path):
        raise FileNotFoundError(f"File TS_03 mancante: {ts3_path}")

    year, month = extract_year_month_from_folder(month_folder)
    num_days = calendar.monthrange(year, month)[1]

    if not satac_path:
        default_satac = os.path.join(month_folder, "SATAC_Meter_Day.xlsx")
        satac_path = default_satac if os.path.exists(default_satac) else ""

    ts1_data = parse_scada_weather_file(ts1_path)
    ts3_data = parse_scada_weather_file(ts3_path)

    if not output_path:
        output_filename = f"{year:04d} {month:02d} Dati Humidita.xlsx"
        output_path = os.path.join(month_folder, output_filename)

    month_en = ENGLISH_MONTHS.get(month, f"Month{month}")
    month_sheet_name = f"{month_en} {year}"
    summary_sheet_name = "Rapporto Umidità"

    wb = openpyxl.Workbook()

    font_family = "Calibri"
    font_title = Font(name=font_family, size=14, bold=True, color="000000")
    font_header_bold = Font(name=font_family, size=11, bold=True, color="000000")
    font_regular = Font(name=font_family, size=11, bold=False, color="000000")
    font_bold = Font(name=font_family, size=11, bold=True, color="000000")

    fill_header_blue = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_light_gray = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    fill_subtle_blue = PatternFill(start_color="EDF2F8", end_color="EDF2F8", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_center_nowrap = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    thin_border_side = Side(style="thin", color="D3D3D3")
    medium_border_side = Side(style="medium", color="000000")

    # =========================================================================
    # SHEET 1: Rapporto Umidità (Summary Report Table)
    # =========================================================================
    ws_summary = wb.active
    ws_summary.title = summary_sheet_name
    ws_summary.views.sheetView[0].showGridLines = True

    # Title: A1:C1 merged
    ws_summary.merge_cells("A1:C1")
    cell_title = ws_summary["A1"]
    cell_title.value = "Rapporto Umidità"
    cell_title.font = font_title
    cell_title.alignment = align_center_nowrap
    ws_summary.row_dimensions[1].height = 24.0

    # Header Row 4 & 5
    ws_summary.row_dimensions[4].height = 20.0
    ws_summary.row_dimensions[5].height = 55.0

    # A4:A5 merged -> "Data"
    ws_summary.merge_cells("A4:A5")
    cell_a4 = ws_summary["A4"]
    cell_a4.value = "Data"
    cell_a4.font = font_header_bold
    cell_a4.alignment = align_center_nowrap
    cell_a4.fill = fill_header_blue

    # B4:C4 merged -> First day date (formatted as mmmm-yy)
    ws_summary.merge_cells("B4:C4")
    cell_b4 = ws_summary["B4"]
    cell_b4.value = datetime(year, month, 1)
    cell_b4.number_format = r"[$-410]mmmm\-yy;@"
    cell_b4.font = font_header_bold
    cell_b4.alignment = align_center_nowrap
    cell_b4.fill = fill_header_blue

    # B5: Relative Humidity header
    cell_b5 = ws_summary["B5"]
    cell_b5.value = "Relative Humidity\n(max/average)\n[%]"
    cell_b5.font = font_regular
    cell_b5.alignment = align_center
    cell_b5.fill = fill_header_blue

    # C5: Absolute Humidity header
    cell_c5 = ws_summary["C5"]
    cell_c5.value = "Absolute Humidity\n(max/average) \n[g/m2]"
    cell_c5.font = font_regular
    cell_c5.alignment = align_center
    cell_c5.fill = fill_header_blue

    for r in (4, 5):
        for c in (1, 2, 3):
            top_s = medium_border_side if r == 4 else None
            bot_s = medium_border_side if r == 5 else thin_border_side
            left_s = medium_border_side if c == 1 else (medium_border_side if c == 2 and r == 5 else thin_border_side)
            right_s = medium_border_side if c == 3 else thin_border_side
            ws_summary.cell(r, c).border = Border(left=left_s, right=right_s, top=top_s, bottom=bot_s)

    ws_summary.column_dimensions["A"].width = 7.0
    ws_summary.column_dimensions["B"].width = 24.0
    ws_summary.column_dimensions["C"].width = 24.0

    # Populate daily rows (Row 6 to 6 + num_days - 1)
    for day in range(1, num_days + 1):
        curr_row = 5 + day
        ws_summary.row_dimensions[curr_row].height = 18.0

        c_a = ws_summary.cell(curr_row, 1)
        c_a.value = day
        c_a.font = font_regular
        c_a.alignment = align_right
        c_a.border = Border(left=medium_border_side, right=thin_border_side,
                            top=thin_border_side, bottom=thin_border_side)

        source_row = 13 + day

        c_b = ws_summary.cell(curr_row, 2)
        c_b.value = f"='{month_sheet_name}'!D{source_row}"
        c_b.number_format = "0.00"
        c_b.font = font_regular
        c_b.alignment = align_right
        c_b.fill = fill_light_gray
        c_b.border = Border(left=thin_border_side, right=thin_border_side,
                            top=thin_border_side, bottom=thin_border_side)

        c_c = ws_summary.cell(curr_row, 3)
        c_c.value = f"='{month_sheet_name}'!G{source_row}"
        c_c.number_format = "0.00"
        c_c.font = font_regular
        c_c.alignment = align_right
        c_c.fill = fill_light_gray
        c_c.border = Border(left=thin_border_side, right=medium_border_side,
                            top=thin_border_side, bottom=thin_border_side)

    # Summary Row: Media
    summary_row = 6 + num_days
    ws_summary.row_dimensions[summary_row].height = 20.0
    month_totals_row = 14 + num_days

    c_media = ws_summary.cell(summary_row, 1)
    c_media.value = "Media"
    c_media.font = font_bold
    c_media.alignment = align_right
    c_media.border = Border(left=medium_border_side, right=medium_border_side,
                            top=medium_border_side, bottom=medium_border_side)

    c_media_rel = ws_summary.cell(summary_row, 2)
    c_media_rel.value = f"='{month_sheet_name}'!D{month_totals_row}"
    c_media_rel.number_format = "0.00"
    c_media_rel.font = font_bold
    c_media_rel.alignment = align_right
    c_media_rel.fill = fill_subtle_blue
    c_media_rel.border = Border(left=medium_border_side, right=thin_border_side,
                                top=medium_border_side, bottom=medium_border_side)

    c_media_abs = ws_summary.cell(summary_row, 3)
    c_media_abs.value = f"='{month_sheet_name}'!G{month_totals_row}"
    c_media_abs.number_format = "0.00"
    c_media_abs.font = font_bold
    c_media_abs.alignment = align_right
    c_media_abs.fill = fill_subtle_blue
    c_media_abs.border = Border(left=thin_border_side, right=medium_border_side,
                                top=medium_border_side, bottom=medium_border_side)

    # =========================================================================
    # SHEET 2: <Month Name> <YYYY> (Detailed SCADA Calculations)
    # =========================================================================
    ws_month = wb.create_sheet(title=month_sheet_name)
    ws_month.views.sheetView[0].showGridLines = True

    ws_month.row_dimensions[1].height = 25.5
    ws_month.row_dimensions[2].height = 20.0
    ws_month.row_dimensions[8].height = 18.0
    ws_month.row_dimensions[12].height = 18.0
    ws_month.row_dimensions[13].height = 70.0

    table_name = f"TabellaHum{year}{month:02d}"

    ws_month.merge_cells("A1:K1")
    ws_month["A1"] = "Guasti Giornalieri"
    ws_month["A1"].font = font_title
    ws_month["A1"].alignment = align_center_nowrap
    ws_month["A1"].fill = fill_header_blue

    ws_month.merge_cells("A2:K2")
    ws_month["A2"] = month_sheet_name
    ws_month["A2"].font = font_title
    ws_month["A2"].alignment = align_center_nowrap
    ws_month["A2"].fill = fill_header_blue

    ws_month.merge_cells("A3:G3")
    ws_month["A3"] = "Total Days"
    ws_month["A3"].font = font_header_bold
    ws_month["A3"].alignment = align_left
    ws_month.merge_cells("I3:K3")
    ws_month["I3"] = f"=COUNT({table_name}[[#All],[Colonna1]])"
    ws_month["I3"].font = font_bold
    ws_month["I3"].alignment = align_center_nowrap

    ws_month.merge_cells("A4:G4")
    ws_month["A4"] = "PVSyst PR for current month"
    ws_month["A4"].font = font_header_bold
    ws_month["A4"].alignment = align_left

    ws_month.merge_cells("A5:G5")
    ws_month["A5"] = "RAW PR"
    ws_month["A5"].font = font_header_bold
    ws_month["A5"].alignment = align_left

    ws_month.merge_cells("A6:G6")
    ws_month["A6"] = "Irradiance Acceptance Rate [%]"
    ws_month["A6"].font = font_header_bold
    ws_month["A6"].alignment = align_left
    ws_month.merge_cells("I6:K6")
    ws_month["I6"] = 0.03
    ws_month["I6"].number_format = "0.00"
    ws_month["I6"].font = font_bold
    ws_month["I6"].alignment = align_center_nowrap

    ws_month.merge_cells("A7:G7")
    ws_month["A7"] = "Humidity Acceptance Rate [%]"
    ws_month["A7"].font = font_header_bold
    ws_month["A7"].alignment = align_left
    ws_month.merge_cells("I7:K7")
    ws_month["I7"] = acceptance_rate
    ws_month["I7"].number_format = "0.00"
    ws_month["I7"].font = font_bold
    ws_month["I7"].alignment = align_center_nowrap

    ws_month.merge_cells("A8:K8")
    ws_month["A8"] = "FILE PATHS"
    ws_month["A8"].font = font_header_bold
    ws_month["A8"].alignment = align_left
    ws_month["A8"].fill = fill_header_blue

    ws_month["A9"] = "SATAC Meter"
    ws_month["A9"].font = font_header_bold
    ws_month.merge_cells("B9:K9")
    ws_month["B9"] = satac_path
    ws_month["B9"].font = font_regular

    ws_month["A10"] = "Weather Data TX1"
    ws_month["A10"].font = font_header_bold
    ws_month.merge_cells("B10:K10")
    ws_month["B10"] = ts1_path
    ws_month["B10"].font = font_regular

    ws_month["A11"] = "Weather Data TX3"
    ws_month["A11"].font = font_header_bold
    ws_month.merge_cells("B11:K11")
    ws_month["B11"] = ts3_path
    ws_month["B11"].font = font_regular

    headers = [
        "Data",
        "Relative Humidity\nTX1\n[%]",
        "Relative Humidity\nTX3\n[%]",
        "Relative Humidity\n(max/average)\n[%]",
        "Absolute Humidity\nTX1\n[g/m2]",
        "Absolute Humidity\nTX3\n[g/m2]",
        "Absolute Humidity\n(max/average) \n[g/m2]"
    ]

    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws_month.cell(13, col_idx)
        cell.value = h_text
        cell.font = font_header_bold
        cell.alignment = align_center
        cell.fill = fill_header_blue

    comb_rel_list = []
    comb_abs_list = []

    for day in range(1, num_days + 1):
        curr_row = 13 + day
        dt_day = datetime(year, month, day)
        d_key = dt_day.date()

        ws_month.row_dimensions[curr_row].height = 18.0

        val_t1_rel = ts1_data.get((d_key, "relative"), 0.0)
        val_t3_rel = ts3_data.get((d_key, "relative"), 0.0)
        val_t1_abs = ts1_data.get((d_key, "absolute"), 0.0)
        val_t3_abs = ts3_data.get((d_key, "absolute"), 0.0)

        comb_rel = compute_combined_humidity(val_t1_rel, val_t3_rel, "relative", acceptance_rate)
        comb_abs = compute_combined_humidity(val_t1_abs, val_t3_abs, "absolute", acceptance_rate)
        if comb_rel > 0:
            comb_rel_list.append(comb_rel)
        if comb_abs > 0:
            comb_abs_list.append(comb_abs)

        # Col A: Date
        c_dt = ws_month.cell(curr_row, 1)
        c_dt.value = dt_day
        c_dt.number_format = r"[$-410]d\-mmm\-yy;@"
        c_dt.font = font_regular
        c_dt.alignment = align_center_nowrap

        # Col B: Relative Humidity TX1 [%]
        c_b = ws_month.cell(curr_row, 2)
        c_b.value = round(val_t1_rel, 3)
        c_b.number_format = "0.000"
        c_b.font = font_regular
        c_b.alignment = align_right

        # Col C: Relative Humidity TX3 [%]
        c_c = ws_month.cell(curr_row, 3)
        c_c.value = round(val_t3_rel, 3)
        c_c.number_format = "0.000"
        c_c.font = font_regular
        c_c.alignment = align_right

        # Col D: Relative Humidity (max/average) [%]
        c_d = ws_month.cell(curr_row, 4)
        c_d.value = f"=IF(ABS(B{curr_row}-C{curr_row})>$I$7*100,MAX(B{curr_row},C{curr_row}),AVERAGE(B{curr_row},C{curr_row}))"
        c_d.number_format = "0.000"
        c_d.font = font_regular
        c_d.alignment = align_right

        # Col E: Absolute Humidity TX1 [g/m2]
        c_e = ws_month.cell(curr_row, 5)
        c_e.value = round(val_t1_abs, 3)
        c_e.number_format = "0.000"
        c_e.font = font_regular
        c_e.alignment = align_right

        # Col F: Absolute Humidity TX3 [g/m2]
        c_f = ws_month.cell(curr_row, 6)
        c_f.value = round(val_t3_abs, 3)
        c_f.number_format = "0.000"
        c_f.font = font_regular
        c_f.alignment = align_right

        # Col G: Absolute Humidity (max/average) [g/m2]
        c_g = ws_month.cell(curr_row, 7)
        c_g.value = f"=IF(ABS(F{curr_row}-E{curr_row})>$I$7*E{curr_row},MAX(E{curr_row},F{curr_row}),AVERAGE(E{curr_row},F{curr_row}))"
        c_g.number_format = "0.000"
        c_g.font = font_regular
        c_g.alignment = align_right

    # Month Totals / Average Row
    last_data_row = 13 + num_days
    totals_row = last_data_row + 1
    ws_month.row_dimensions[totals_row].height = 20.0

    c_tot_rel = ws_month.cell(totals_row, 4)
    c_tot_rel.value = f'=AVERAGEIF(D14:D{last_data_row}, ">0")'
    c_tot_rel.number_format = "0.000"
    c_tot_rel.font = font_bold
    c_tot_rel.alignment = align_right

    c_tot_abs = ws_month.cell(totals_row, 7)
    c_tot_abs.value = f'=AVERAGEIF(G14:G{last_data_row}, ">0")'
    c_tot_abs.number_format = "0.000"
    c_tot_abs.font = font_bold
    c_tot_abs.alignment = align_right

    tab_ref = f"A13:G{last_data_row}"
    table = Table(displayName=table_name, ref=tab_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    ws_month.add_table(table)

    ws_month.column_dimensions["A"].width = 14.0
    ws_month.column_dimensions["B"].width = 16.0
    ws_month.column_dimensions["C"].width = 16.0
    ws_month.column_dimensions["D"].width = 18.0
    ws_month.column_dimensions["E"].width = 16.0
    ws_month.column_dimensions["F"].width = 16.0
    ws_month.column_dimensions["G"].width = 18.0
    ws_month.column_dimensions["H"].width = 6.0
    ws_month.column_dimensions["I"].width = 16.0

    output_path = os.path.abspath(output_path)
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    wb.save(output_path)
    wb.close()

    avg_rel = sum(comb_rel_list) / len(comb_rel_list) if comb_rel_list else 0.0
    avg_abs = sum(comb_abs_list) / len(comb_abs_list) if comb_abs_list else 0.0

    return {
        "success": True,
        "year": year,
        "month": month,
        "month_name": month_sheet_name,
        "num_days": num_days,
        "operational_days": len(comb_rel_list),
        "avg_rel_humidity": avg_rel,
        "avg_abs_humidity": avg_abs,
        "output_path": output_path
    }


def find_all_month_folders(base_dir=DEFAULT_BASE_DIR):
    """Scan base_dir for month folders containing both TS_01 and TS_03."""
    if not os.path.exists(base_dir):
        return []

    month_folders = []
    for item in sorted(os.listdir(base_dir)):
        item_path = os.path.join(base_dir, item)
        if os.path.isdir(item_path):
            ts1 = os.path.join(item_path, "TS_01_Weather_1Day.xlsx")
            ts3 = os.path.join(item_path, "TS_03_Weather_1Day.xlsx")
            if os.path.exists(ts1) and os.path.exists(ts3):
                month_folders.append(item_path)
    return month_folders


def generate_all_reports(base_dir=DEFAULT_BASE_DIR, log_callback=print):
    """Process all available month folders in base_dir."""
    folders = find_all_month_folders(base_dir)
    if not folders:
        log_callback(f"[AVVISO] Nessuna cartella mensile con file SCADA trovata in: {base_dir}")
        return []

    results = []
    log_callback(f"[INFO] Trovate {len(folders)} cartelle mensili con dati SCADA. Inizio elaborazione...")

    for fpath in folders:
        folder_name = os.path.basename(fpath)
        log_callback(f"\n---> Elaborazione {folder_name}...")
        try:
            res = generate_humidity_report(fpath)
            log_callback(f"     [OK] Generato: {os.path.basename(res['output_path'])}")
            log_callback(f"          Giorni con dati: {res['operational_days']}/{res['num_days']}")
            log_callback(f"          Media Umidità Relativa: {res['avg_rel_humidity']:.2f}%")
            log_callback(f"          Media Umidità Assoluta: {res['avg_abs_humidity']:.2f} g/m³")
            results.append(res)
        except Exception as e:
            log_callback(f"     [ERRORE] Impossibile elaborare {folder_name}: {e}")

    log_callback(f"\n[FINE] Completata elaborazione: {len(results)}/{len(folders)} report generati con successo.")
    return results


def launch_gui():
    """Launch modern Tkinter desktop GUI for the Humidity Report Generator."""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import threading
    import subprocess

    root = tk.Tk()
    root.title("Generatore Report Umidità SCADA - Mazara Solar")
    root.geometry("800x650")
    root.minsize(700, 540)

    # Windows App ID for taskbar icon
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("get.mazara.humidityreport.1.0")
    except Exception:
        pass

    # Set Window and Taskbar Icon
    icon_ico = get_resource_path(os.path.join("assets", "logo_umidita.ico"))
    if os.path.exists(icon_ico):
        try:
            root.iconbitmap(icon_ico)
        except Exception:
            pass

    icon_png = get_resource_path(os.path.join("assets", "logo_umidita.png"))
    logo_img_ref = [None]
    if os.path.exists(icon_png):
        try:
            from PIL import Image, ImageTk
            pil_img = Image.open(icon_png).resize((50, 50), Image.Resampling.LANCZOS)
            logo_img_ref[0] = ImageTk.PhotoImage(pil_img)
            root.iconphoto(True, logo_img_ref[0])
        except Exception:
            try:
                logo_img_ref[0] = tk.PhotoImage(file=icon_png)
                root.iconphoto(True, logo_img_ref[0])
            except Exception:
                pass

    BG_COLOR = "#F7F9FC"
    CARD_BG = "#FFFFFF"
    PRIMARY_COLOR = "#1B365D"
    ACCENT_COLOR = "#0078D4"
    SUCCESS_COLOR = "#107C41"

    root.configure(bg=BG_COLOR)

    header_frame = tk.Frame(root, bg=PRIMARY_COLOR, padx=16, pady=12)
    header_frame.pack(fill="x", side="top")

    # Logo in header
    if logo_img_ref[0]:
        lbl_logo = tk.Label(header_frame, image=logo_img_ref[0], bg=PRIMARY_COLOR)
        lbl_logo.pack(side="left", padx=(0, 14))

    title_box = tk.Frame(header_frame, bg=PRIMARY_COLOR)
    title_box.pack(side="left", fill="both", expand=True)

    lbl_title = tk.Label(
        title_box,
        text="SCADA Humidity Report Generator",
        font=("Segoe UI", 16, "bold"),
        fg="#FFFFFF",
        bg=PRIMARY_COLOR
    )
    lbl_title.pack(anchor="w")

    lbl_sub = tk.Label(
        title_box,
        text="Generazione automatica del report di umidità mensile da TS_01 e TS_03 (Mazara 01 A2A)",
        font=("Segoe UI", 9),
        fg="#D9E1F2",
        bg=PRIMARY_COLOR
    )
    lbl_sub.pack(anchor="w", pady=(2, 0))

    main_container = tk.Frame(root, bg=BG_COLOR, padx=16, pady=12)
    main_container.pack(fill="both", expand=True)

    card_config = tk.LabelFrame(
        main_container,
        text=" Selezione Dati SCADA Mensili ",
        font=("Segoe UI", 10, "bold"),
        fg=PRIMARY_COLOR,
        bg=CARD_BG,
        relief="groove",
        bd=1,
        padx=14,
        pady=10
    )
    card_config.pack(fill="x", pady=(0, 10))

    frame_dir = tk.Frame(card_config, bg=CARD_BG)
    frame_dir.pack(fill="x", pady=4)
    tk.Label(frame_dir, text="Cartella Base:", font=("Segoe UI", 9, "bold"), bg=CARD_BG, width=14, anchor="w").pack(side="left")
    base_dir_var = tk.StringVar(value=DEFAULT_BASE_DIR)
    entry_base_dir = tk.Entry(frame_dir, textvariable=base_dir_var, font=("Segoe UI", 9), relief="solid", bd=1)
    entry_base_dir.pack(side="left", fill="x", expand=True, padx=(4, 6))

    def browse_base_dir():
        d = filedialog.askdirectory(initialdir=base_dir_var.get(), title="Seleziona Cartella Base Daily Reports")
        if d:
            base_dir_var.set(d)
            refresh_months()

    btn_browse_base = tk.Button(frame_dir, text="Sfoglia...", font=("Segoe UI", 9), command=browse_base_dir, relief="groove")
    btn_browse_base.pack(side="right")

    frame_month = tk.Frame(card_config, bg=CARD_BG)
    frame_month.pack(fill="x", pady=6)
    tk.Label(frame_month, text="Mese SCADA:", font=("Segoe UI", 9, "bold"), bg=CARD_BG, width=14, anchor="w").pack(side="left")

    month_combo = ttk.Combobox(frame_month, font=("Segoe UI", 9), state="readonly", width=25)
    month_combo.pack(side="left", padx=(4, 8))

    btn_refresh = tk.Button(frame_month, text="🔄 Ricarica Mesi", font=("Segoe UI", 9), relief="groove")
    btn_refresh.pack(side="left", padx=(0, 12))

    lbl_detected = tk.Label(frame_month, text="", font=("Segoe UI", 9, "italic"), fg="#555555", bg=CARD_BG)
    lbl_detected.pack(side="left")

    frame_custom = tk.Frame(card_config, bg=CARD_BG)
    frame_custom.pack(fill="x", pady=4)
    tk.Label(frame_custom, text="Oppure Cartella:", font=("Segoe UI", 9, "bold"), bg=CARD_BG, width=14, anchor="w").pack(side="left")
    custom_folder_var = tk.StringVar(value="")
    entry_custom = tk.Entry(frame_custom, textvariable=custom_folder_var, font=("Segoe UI", 9), relief="solid", bd=1)
    entry_custom.pack(side="left", fill="x", expand=True, padx=(4, 6))

    def browse_custom_folder():
        d = filedialog.askdirectory(initialdir=base_dir_var.get(), title="Seleziona Cartella Mensile Specifica")
        if d:
            custom_folder_var.set(d)

    btn_browse_custom = tk.Button(frame_custom, text="Sfoglia...", font=("Segoe UI", 9), command=browse_custom_folder, relief="groove")
    btn_browse_custom.pack(side="right")

    frame_actions = tk.Frame(main_container, bg=BG_COLOR)
    frame_actions.pack(fill="x", pady=(0, 10))

    last_output_dir = [None]

    def log(msg):
        text_log.configure(state="normal")
        text_log.insert("end", str(msg) + "\n")
        text_log.see("end")
        text_log.configure(state="disabled")

    def run_generate_single():
        custom_p = custom_folder_var.get().strip()
        if custom_p:
            target_folder = custom_p
        else:
            selected_m = month_combo.get().strip()
            if not selected_m:
                messagebox.showwarning("Attenzione", "Seleziona un mese dalla lista o specifica una cartella personalizzata.")
                return
            target_folder = os.path.join(base_dir_var.get(), selected_m)

        def worker():
            btn_gen_one.config(state="disabled")
            btn_gen_all.config(state="disabled")
            log(f"\n==================================================")
            log(f"Avvio generazione report per: {os.path.basename(target_folder)}")
            log(f"Percorso: {target_folder}")
            try:
                res = generate_humidity_report(target_folder)
                last_output_dir[0] = os.path.dirname(res["output_path"])
                log(f"[SUCCESSO] File creato con successo!")
                log(f"  File: {res['output_path']}")
                log(f"  Giorni con dati: {res['operational_days']}/{res['num_days']}")
                log(f"  Umidità Relativa Media: {res['avg_rel_humidity']:.2f}%")
                log(f"  Umidità Assoluta Media: {res['avg_abs_humidity']:.2f} g/m³")
                btn_open_folder.config(state="normal")
                messagebox.showinfo("Completato", f"Report Umidità generato con successo:\n{os.path.basename(res['output_path'])}")
            except Exception as e:
                log(f"[ERRORE] Generazione fallita: {e}")
                messagebox.showerror("Errore", f"Errore durante la generazione:\n{e}")
            finally:
                btn_gen_one.config(state="normal")
                btn_gen_all.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    def run_generate_all():
        b_dir = base_dir_var.get().strip()
        if not os.path.exists(b_dir):
            messagebox.showerror("Errore", f"Cartella base non trovata: {b_dir}")
            return

        confirm = messagebox.askyesno(
            "Conferma Batch",
            "Vuoi elaborare tutti i mesi disponibili con dati SCADA nella cartella base?"
        )
        if not confirm:
            return

        def worker():
            btn_gen_one.config(state="disabled")
            btn_gen_all.config(state="disabled")
            log(f"\n==================================================")
            log(f"Avvio elaborazione BATCH per tutti i mesi in:\n{b_dir}")
            try:
                results = generate_all_reports(b_dir, log_callback=log)
                if results:
                    last_output_dir[0] = os.path.dirname(results[-1]["output_path"])
                    btn_open_folder.config(state="normal")
                    messagebox.showinfo("Completato", f"Elaborazione batch completata!\n{len(results)} report generati.")
            except Exception as e:
                log(f"[ERRORE BATCH] {e}")
                messagebox.showerror("Errore", f"Errore durante l'elaborazione batch:\n{e}")
            finally:
                btn_gen_one.config(state="normal")
                btn_gen_all.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    def open_output_folder():
        p = last_output_dir[0] or base_dir_var.get()
        if os.path.exists(p):
            if sys.platform == "win32":
                os.startfile(p)
            else:
                subprocess.Popen(["xdg-open", p])

    btn_gen_one = tk.Button(
        frame_actions,
        text="⚡ Genera Report Mese Selezionato",
        font=("Segoe UI", 10, "bold"),
        bg=ACCENT_COLOR,
        fg="#FFFFFF",
        activebackground="#005A9E",
        activeforeground="#FFFFFF",
        relief="groove",
        padx=14,
        pady=6,
        command=run_generate_single
    )
    btn_gen_one.pack(side="left", padx=(0, 8))

    btn_gen_all = tk.Button(
        frame_actions,
        text="📁 Genera per TUTTI i Mesi Disponibili",
        font=("Segoe UI", 10, "bold"),
        bg=SUCCESS_COLOR,
        fg="#FFFFFF",
        activebackground="#0E6B37",
        activeforeground="#FFFFFF",
        relief="groove",
        padx=14,
        pady=6,
        command=run_generate_all
    )
    btn_gen_all.pack(side="left", padx=(0, 8))

    btn_open_folder = tk.Button(
        frame_actions,
        text="📂 Apri Cartella",
        font=("Segoe UI", 9),
        state="disabled",
        relief="groove",
        padx=10,
        pady=6,
        command=open_output_folder
    )
    btn_open_folder.pack(side="right")

    card_log = tk.LabelFrame(
        main_container,
        text=" Log Operazioni ",
        font=("Segoe UI", 10, "bold"),
        fg=PRIMARY_COLOR,
        bg=CARD_BG,
        relief="groove",
        bd=1,
        padx=10,
        pady=8
    )
    card_log.pack(fill="both", expand=True)

    text_log = tk.Text(card_log, font=("Consolas", 9), bg="#1E1E1E", fg="#D4D4D4", insertbackground="#FFFFFF", relief="flat")
    scroll_y = tk.Scrollbar(card_log, orient="vertical", command=text_log.yview)
    text_log.configure(yscrollcommand=scroll_y.set)
    scroll_y.pack(side="right", fill="y")
    text_log.pack(side="left", fill="both", expand=True)
    text_log.configure(state="disabled")

    def refresh_months():
        b_dir = base_dir_var.get().strip()
        month_combo["values"] = []
        lbl_detected.config(text="Ricerca cartelle...")
        root.update_idletasks()

        folders = find_all_month_folders(b_dir)
        names = [os.path.basename(f) for f in folders]
        month_combo["values"] = names
        if names:
            prev_m = get_previous_month_str()
            if prev_m in names:
                month_combo.set(prev_m)
            else:
                month_combo.set(names[-1])
            lbl_detected.config(text=f"{len(names)} mesi disponibili (default: {month_combo.get()})", fg="#107C41")
        else:
            month_combo.set("")
            lbl_detected.config(text="Nessun mese trovato con TS_01/03", fg="#D83B01")

    btn_refresh.config(command=refresh_months)
    root.after(100, refresh_months)
    root.mainloop()


def main():
    default_prev_month = get_previous_month_str()
    parser = argparse.ArgumentParser(
        description="SCADA Humidity Report Generator - Generates Excel monthly humidity reports from TS_01 and TS_03 weather data."
    )
    parser.add_argument("--month", type=str, nargs="?", const=default_prev_month,
                        help=f"Month to process (e.g. '2026 08'). Defaults to previous month ({default_prev_month}) if passed without value.")
    parser.add_argument("--folder", type=str, help="Explicit path to a month folder containing TS_01 and TS_03 files")
    parser.add_argument("--output", type=str, help="Explicit output .xlsx file path")
    parser.add_argument("--all", action="store_true", help="Process all available month folders in base directory")
    parser.add_argument("--base-dir", type=str, default=DEFAULT_BASE_DIR, help="Base directory for Daily Reports")
    parser.add_argument("--rate", type=float, default=0.05, help="Humidity acceptance rate (default: 0.05 = 5%%)")
    parser.add_argument("--gui", action="store_true", help="Launch interactive graphical interface")

    args = parser.parse_args()

    if args.gui or (len(sys.argv) == 1 and ("DISPLAY" in os.environ or sys.platform == "win32")):
        try:
            launch_gui()
            return
        except Exception as e:
            print(f"Impossibile avviare la GUI ({e}). Passaggio alla modalità da riga di comando.")

    if args.all:
        generate_all_reports(args.base_dir, log_callback=print)
    elif args.folder:
        res = generate_humidity_report(args.folder, output_path=args.output, acceptance_rate=args.rate)
        print(f"[OK] Report generato con successo:")
        print(f"     File: {res['output_path']}")
        print(f"     Giorni con dati: {res['operational_days']}/{res['num_days']}")
        print(f"     Media Umidità Relativa: {res['avg_rel_humidity']:.2f}%")
        print(f"     Media Umidità Assoluta: {res['avg_abs_humidity']:.2f} g/m³")
    elif args.month:
        month_folder = os.path.join(args.base_dir, args.month)
        res = generate_humidity_report(month_folder, output_path=args.output, acceptance_rate=args.rate)
        print(f"[OK] Report generato con successo per {args.month}:")
        print(f"     File: {res['output_path']}")
        print(f"     Giorni con dati: {res['operational_days']}/{res['num_days']}")
        print(f"     Media Umidità Relativa: {res['avg_rel_humidity']:.2f}%")
        print(f"     Media Umidità Assoluta: {res['avg_abs_humidity']:.2f} g/m³")
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
