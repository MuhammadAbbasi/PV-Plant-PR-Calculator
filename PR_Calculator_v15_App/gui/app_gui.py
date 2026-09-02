import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from config.plant_config import PLANT_DC_POWER_TOTAL, PLANT_AC_POWER_TOTAL, PVSYST_MONTHLY_PR
from config.constants import ITALIAN_MONTHS_FULL
from gui.components.log_viewer import LogRedirector
from data_converters.vcom_converter import convert_vcom_to_scada
from excel_engine.mother_file_syncer import sync_mother_file

class PRCalculatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PR Calculator v15.0 - Professional Solar Analytics")
        self.root.geometry("900x650")
        
        # Setup UI Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tab_calc = ttk.Frame(self.notebook)
        self.tab_vcom = ttk.Frame(self.notebook)
        self.tab_logs = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_calc, text="PR Recalculation")
        self.notebook.add(self.tab_vcom, text="VCOM Data Converter")
        self.notebook.add(self.tab_logs, text="System Logs")
        
        self._build_calc_tab()
        self._build_vcom_tab()
        self._build_logs_tab()

    def _build_calc_tab(self):
        lbl = ttk.Label(self.tab_calc, text="PR Recalculation Engine v15.0", font=("Segoe UI", 14, "bold"))
        lbl.pack(anchor=tk.W, padx=15, pady=15)
        
        btn_frame = ttk.Frame(self.tab_calc)
        btn_frame.pack(fill=tk.X, padx=15, pady=10)
        
        ttk.Button(btn_frame, text="Run Single Day PR", command=self._run_single_day).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Sync Mother File", command=self._sync_mother).pack(side=tk.LEFT, padx=5)

    def _build_vcom_tab(self):
        lbl = ttk.Label(self.tab_vcom, text="VCOM 3-File to SCADA Converter", font=("Segoe UI", 14, "bold"))
        lbl.pack(anchor=tk.W, padx=15, pady=15)

    def _build_logs_tab(self):
        txt_log = tk.Text(self.tab_logs, wrap=tk.WORD, font=("Consolas", 10))
        txt_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        sys.stdout = LogRedirector(txt_log)
        print(">>> PR Calculator v15.0 System initialized.")

    def _run_single_day(self):
        messagebox.showinfo("PR Calculator v15", "Single Day Calculation Module Ready!")

    def _sync_mother(self):
        messagebox.showinfo("PR Calculator v15", "Mother File Synchronization Module Ready!")
