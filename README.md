# PV Plant Performance Ratio (PR) Calculator - Mazara 01

A professional, high-performance Python-based tool designed for the **GET S.R.L.** Mazara 01 photovoltaic plant. This application automates the calculation of the Performance Ratio (PR), providing both raw and compensated metrics by processing SCADA data and weather station logs.

![GUI Preview](gui_annotated_guide.png)

## 🚀 Features

- **Automated Calculation Engine**: Processes 15-minute interval data for active power, solar irradiance (POA), and energy meter readings.
- **Compensated PR Analysis**: Intelligent logic to account for:
    - **Curtailment Losses**: Energy lost due to grid-imposed power limits.
    - **Downtime Losses**: Energy lost during inverter or transformer outages.
- **Batch Processing Mode**: Quickly process an entire month's worth of data by selecting a parent folder containing daily subdirectories.
- **Excel Automation**: Utilizes Excel COM (ActiveX) for seamless report generation, ensuring that complex formulas and templates remain uncorrupted.
- **Mother-Child File Syncing**: Automatically updates monthly "Mother" files with data from daily "Child" recalculation files.
- **Modern Dark Mode Interface**: A premium, obsidian-themed GUI built with Tkinter, featuring real-time logging and performance metrics.

## 📋 Prerequisites

To run the source code, you need Python 3.8+ and the following dependencies:

- **Windows OS**: Required for Excel COM integration.
- **Microsoft Excel**: Installed on the machine.
- **Python Libraries**:
  ```bash
  pip install pandas numpy openpyxl pywin32 Pillow
  ```

## 🛠️ Installation & Setup

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

## 📖 How to Use

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

## 🧮 Calculation Logic

The tool follows a rigorous technical methodology:
- **Irradiance Thresholding**: Only calculates PR when average POA is above a configurable threshold (default 50 W/m²).
- **Compensated PR**: 
  $PR_{comp} = \frac{E_{gen} + E_{loss\_dt} + E_{loss\_curt}}{P_{nom} \times \frac{H}{G_{STC}}}$
- **Energy Meter Delta**: Calculates real energy production from cumulative active energy readings (SATAC meter).

## 📄 Documentation

For a detailed walkthrough, refer to the included [Manuale_Utente_PR_Calculator.html](Manuale_Utente_PR_Calculator.html).

## 🤝 Developed By

**Muhammad Abbasi**

Data Scientist and Automation Engineer - GET S.R.L.

---
*Note: This tool is specifically tailored for the Mazara 01 plant configuration but can be adapted for other PV infrastructures.*
