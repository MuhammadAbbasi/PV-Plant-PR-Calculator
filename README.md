# PV Plant Performance Ratio (PR) Calculator - Mazara 01 (v5.0)

A professional, high-performance Python-based tool designed for the **GET S.R.L.** Mazara 01 photovoltaic plant. This application automates the calculation of the Performance Ratio (PR), providing both raw and compensated metrics by processing SCADA data and weather station logs.

![GUI Preview](gui_annotated_guide.png)

## Features

- **Automated Calculation Engine**: Processes 15-minute interval data for active power, solar irradiance (POA), and energy meter readings.
- **Compensated PR Analysis**: Intelligent logic to account for:
    - **Curtailment Losses**: Energy lost due to grid-imposed power limits.
    - **Downtime Losses**: Energy lost during inverter or transformer outages.
- **Batch Processing Mode**: Quickly process an entire month's worth of data by selecting a parent folder containing daily subdirectories.
- **Excel Automation via ActiveX**: Utilizes Excel COM for seamless report generation, avoiding `openpyxl` table corruption and ensuring that complex formulas and styles remain uncorrupted.
- **Mother-Child File Syncing (v5.0 aligned)**: Automatically links monthly "Mother" files with data from daily "Child" recalculation files via dynamic Excel formulas.
- **Obsidian Dark Mode Interface**: A premium, luxury-themed GUI built with Tkinter, featuring real-time logging, interactive controls, and performance metrics.

## Prerequisites

To run the source code, you need Windows (for Excel COM integration), Microsoft Excel installed, and Python 3.8+ with the following libraries:

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
   - `PR_recalculation_26_apr.xlsx` (Pristine daily template - *Version 5.0 aligned*)

3. **Assets**:
   Place company logos in the `assets/` folder (`logo.png`, `logo.ico`).

## How to Use

1. **Launch the Application**:
   Run the GUI using Python:
   ```bash
   python PR_Calculator_GUI.py
   ```
   Or run the compiled executable:
   ```bash
   "PR Calculator v5.exe"
   ```

2. **Single Day Processing**:
   - Select the folder containing the SCADA files for the specific day.
   - Enter the target date (`YYYY-MM-DD`).
   - Click **"Calcola Performance Ratio"**.

3. **Batch Processing (Monthly)**:
   - Select a parent folder containing subfolders named by day (e.g., `01`, `02`, `03` ... `31`).
   - Check **"Ricalcola forzatamente i giorni già elaborati"** to overwrite existing daily workbooks.
   - The tool will iterate through every day, generate individual child workbooks, and sync them to the monthly Mother file.

---

## 📊 Excel Templates & Formatting Requirements (v5.0)

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
        *   Column `N`: Active Power Regulation Limit Ratio (expressed as a decimal, e.g., `0.876`)
    *   **Formula Cells (Rows 15 to 110)**: Autopopulated with exact Excel formulas to prevent `#DIV/0!` errors:
        *   Column `G` (POA Avg): `=IFERROR((D{r}+F{r})/2, 0)`
        *   Column `H` (POA threshold check): `=IFERROR(IF(AVERAGE(C{r},E{r})>$BA$7,AVERAGE(C{r},E{r}),0), 0)`
        *   Column `I` (POA max check): `=IFERROR(IF(AND(D{r}=0,F{r}=0),0,IF(H{r}>$BA$6,MAX(D{r},F{r}),G{r})), 0)`
        *   Column `J` (POA difference ratio): `=IFERROR(IF(AND(D{r}>0,F{r}>0),ABS(D{r}-F{r})/AVERAGE(D{r},F{r}),0), 0)`
        *   Column `M` (Active Energy production): `=IFERROR((L{r}-K{r})*1000, 0)`
    *   **Nominal Parameters (Column BA)**:
        *   `BA4`: PVSyst PR Target as decimal (e.g., `0.897` written from GUI)
        *   `BA6`: Irradiance acceptance limit ratio (e.g., `0.03`)
        *   `BA7`: Minimum Irradiance threshold (e.g., `50` written from GUI)
    *   **Shifted Parameters Table (Columns BD & BH)**:
        *   `BD2`: English PR title header (e.g., `"1 May 2026 PR Calculation"`)
        *   `BH3`: Total values count (`=+BA3`)
        *   `BH4`: Total values with POA > 0 (`=COUNTIF(H15:H110,">0")`)
        *   `BH5`: PVSyst PR for current month in % (`=+BA4*100`)
        *   `BH6`: RAW PR in % (`=+BA5*100`)
        *   `BH7`: Average of each PR in % (`=AVERAGE(...)*100`)
        *   `BH8`: **PR from SCADA (Uncompensated PR % written from GUI, e.g., `37.529`)**
        *   `BH9`: Irradiance acceptance limit ratio (`=+BA6*100`)
        *   `BH10`: Minimum Irradiance threshold (`=+BA7`)

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
    *   Column `D` (PR Daily): `='[ChildPath]PR_Calc'!$BA$5*100` (or `='[ChildPath]PR_Calc'!$BH$6`)
    *   Column `E` (PR SCADA): **`='[ChildPath]PR_Calc'!$BH$8`** (Linked to daily SCADA PR in Column BH)
    *   Column `F`, `G`, `H` (Energy Loss): Linked to daily file cells `$AA$111`, `$AN$111`, `$BA$111` respectively.
    *   Columns `I` to `T` (TX1 Inverters): Linked to child file row 111, columns `O` to `Z` (`O$111` to `Z$111`).
    *   Columns `U` to `AF` (TX2 Inverters): Linked to child file row 111, columns `AB` to `AM` (`AB$111` to `AM$111`).
    *   Columns `AG` to `AR` (TX3 Inverters): Linked to child file row 111, columns `AO` to `AZ` (`AO$111` to `AZ$111`).
*   **Summary Row**: Dynamically positioned at Row `5 + num_days` containing the `=AVERAGE(D5:D{last_day_row})` formula in columns `D` and `E`.

---

## 🛠 Compilation (Building the Executable)

To bundle the application into a standalone Windows executable:

```powershell
python scratch/build_exe.py
```

This script will verify your PyInstaller installation, build the executable using **`PR Calculator v5.spec`**, and copy the standalone **`PR Calculator v5.exe`** to the main folder.

---

## Developed By

**Muhammad Abbasi**  
*Data Scientist and Automation Engineer - GET S.R.L.*

---
*Note: This tool is specifically tailored for the Mazara 01 plant configuration but can be adapted for other PV infrastructures.*
