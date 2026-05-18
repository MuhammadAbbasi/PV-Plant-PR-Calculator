# PV Plant Performance Ratio (PR) Calculator - Mazara 01

A professional, high-performance Python-based tool designed for the **GET S.R.L.** Mazara 01 photovoltaic plant. This application automates the calculation of the Performance Ratio (PR), providing both raw and compensated metrics by processing SCADA data and weather station logs.

![GUI Preview](gui_annotated_guide.png)

## Features

- **Automated Calculation Engine**: Processes 15-minute interval data for active power, solar irradiance (POA), and energy meter readings.
- **Compensated PR Analysis**: Intelligent logic to account for:
    - **Curtailment Losses**: Energy lost due to grid-imposed power limits.
    - **Downtime Losses**: Energy lost during inverter or transformer outages.
- **Batch Processing Mode**: Quickly process an entire month's worth of data by selecting a parent folder containing daily subdirectories.
- **Excel Automation**: Utilizes Excel COM (ActiveX) for seamless report generation, ensuring that complex formulas and templates remain uncorrupted.
- **Mother-Child File Syncing**: Automatically updates monthly "Mother" files with data from daily "Child" recalculation files.
- **Modern Dark Mode Interface**: A premium, obsidian-themed GUI built with Tkinter, featuring real-time logging and performance metrics.

## Prerequisites

To run the source code, you need Python 3.8+ and the following dependencies:

- **Windows OS**: Required for Excel COM integration.
- **Microsoft Excel**: Installed on the machine.
- **Python Libraries**:
  ```bash
  pip install pandas numpy openpyxl pywin32 Pillow
  ```

## Installation & Setup

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/MuhammadAbbasi/PV-Plant-PR-Calculator.git
   cd PV-Plant-PR-Calculator
   ```

2. **Template Configuration**:
   Ensure the `original_format/` directory contains the required Excel templates:
   - `00 PR_recalculation_*.xlsx` (Monthly Mother file)
   - `PR_recalculation_*.xlsx` (Daily template)

3. **Assets**:
   Place company logos in the `assets/` folder (`logo.png`, `logo.ico`).

## How to Use

1. **Launch the Application**:
   Run the script using Python:
   ```bash
   python PR_Calculator_GUI.py
   ```

2. **Single Day Processing**:
   - Select the folder containing the SCADA files for the specific day.
   - Enter the target date (YYYY-MM-DD).
   - Click **"Calcola Performance Ratio"**.

3. **Batch Processing (Monthly)**:
   - Select a parent folder containing subfolders named by day (e.g., `01`, `02`, `03`...).
   - Check **"Ricalcola forzatamente i giorni già elaborati"** if you wish to overwrite existing reports.
   - The tool will iterate through every day, generate individual reports, and sync them to the monthly Mother file.

## 📊 Excel Templates & Formatting Requirements

The tool automates calculations by reading from and writing to specific sheets, columns, and cells within two template types. Below are the formatting requirements to ensure compatibility:

### 1. Daily "Child" Workbook (`PR_recalculation_*.xlsx`)
Must contain at least two worksheets with the following exact names and structure:

*   **`PR_Calc` Sheet**:
    *   **Row 14**: Headers row. Columns `AA`, `AB`, and `AC` are dynamically set to:
        *   `AA14`: `"Energy Loss for TX1\nkW/H"`
        *   `AB14`: `"Inverter status TX2-INV-1"`
        *   `AC14`: `"Inverter status TX2-INV-2"`
    *   **Rows 15 to 110 (96 intervals of 15 min)**:
        *   Column `A`: Date (`YYYY-MM-DD`)
        *   Column `B`: Time slot (`HH:MM:SS` from `00:00:00` to `23:45:00`)
        *   Column `C` & `D`: POA1 (W/m²) and POA1 (kWh/m²)
        *   Column `E` & `F`: POA3 (W/m²) and POA3 (kWh/m²)
        *   Column `K`: Meter Previous Reading (SATAC)
        *   Column `L`: Meter Current Reading (SATAC)
        *   Column `N`: Active Power Regulation Limit Ratio (expressed as a decimal, e.g., 0.876)
    *   **Formula Cells (Rows 15 to 110)**: Autopopulated with exact Excel formulas to prevent `#DIV/0!` errors:
        *   Column `G` (POA Avg): `=IFERROR((D{r}+F{r})/2, 0)`
        *   Column `H` (POA threshold check): `=IFERROR(IF(AVERAGE(C{r},E{r})>$BN$7,AVERAGE(C{r},E{r}),0), 0)`
        *   Column `I` (POA max check): `=IFERROR(IF(AND(D{r}=0,F{r}=0),0,IF(H{r}>$BN$6,MAX(D{r},F{r}),G{r})), 0)`
        *   Column `J` (POA difference ratio): `=IFERROR(IF(AND(D{r}>0,F{r}>0),ABS(D{r}-F{r})/AVERAGE(D{r},F{r}),0), 0)`
        *   Column `M` (Active Energy production): `=IFERROR((L{r}-K{r})*1000, 0)`
    *   **Key Cells**:
        *   `BN4` (Column `BN`): Monthly PVSyst target PR (written from GUI)
        *   `BN7` (Column `BN`): Minimum Irradiance threshold (written from GUI)
        *   `BY8` (Column `BY`): Calculated Uncompensated PR value (written from GUI)

*   **`Inverter_data` Sheet**:
    *   **Rows 15 to 110**:
        *   Column `A`: Date (`YYYY-MM-DD`)
        *   Columns `C` to `N` (12 columns): Active power for **TX1-INV-1** to **TX1-INV-12** (kW)
        *   Columns `R` to `AC` (12 columns): Active power for **TX2-INV-1** to **TX2-INV-12** (kW)
        *   Columns `AG` to `AR` (12 columns): Active power for **TX3-INV-1** to **TX3-INV-12** (kW)

### 2. Monthly "Mother" Workbook (`00 PR_recalculation_*.xlsx`)
Must contain a **`PR_Calc`** sheet structured as follows:
*   **Column `A`**: Date sequence for the entire month (Rows 5 to `5 + num_days - 1`).
*   **Columns `D` to `AR` (Rows 5 to `5 + num_days - 1`)**: Auto-linked formulas referencing the corresponding child workbook:
    *   Column `D` (Daily PR): `='[ChildPath]PR_Calc'!$BA$5*100`
    *   Column `E` (Total PR): `='[ChildPath]PR_Calc'!$BN$5*100`
    *   Column `F`, `G`, `H` (Energy Loss): Linked to daily file cells `$AA$111`, `$AN$111`, `$BA$111` respectively.
    *   Columns `I` to `T` (TX1 Inverters): Linked to child file row 111, columns `O` to `Z` (`O$111` to `Z$111`).
    *   Columns `U` to `AF` (TX2 Inverters): Linked to child file row 111, columns `AB` to `AM` (`AB$111` to `AM$111`).
    *   Columns `AG` to `AR` (TX3 Inverters): Linked to child file row 111, columns `AO` to `AZ` (`AO$111` to `AZ$111`).
*   **Summary Row**: Dynamically repositioned at Row `5 + num_days` containing the `=AVERAGE(D5:D{last_day_row})` formula in columns `D` and `E`.

## Calculation Logic

The tool follows a rigorous technical methodology:
- **Irradiance Thresholding**: Only calculates PR when average POA is above a configurable threshold (default 50 W/m²).
- **Compensated PR**: 
  $PR_{comp} = \frac{E_{gen} + E_{loss\_dt} + E_{loss\_curt}}{P_{nom} \times \frac{H}{G_{STC}}}$
- **Energy Meter Delta**: Calculates real energy production from cumulative active energy readings (SATAC meter).

## Documentation

For a detailed walkthrough, refer to the included [Manuale_Utente_PR_Calculator.html](Manuale_Utente_PR_Calculator.html).

## Developed By

**Muhammad Abbasi**

Data Scientist and Automation Engineer - GET S.R.L.

---
*Note: This tool is specifically tailored for the Mazara 01 plant configuration but can be adapted for other PV infrastructures.*
