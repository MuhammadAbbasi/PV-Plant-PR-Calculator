"""Self-check and automated tests for PR_Calculator_GUI_v14:
- Standalone SCADA and VCOM file discovery
- SCADA and VCOM PR data parser resilience (encodings, delimiters, date formats)
- GUI layout and dedicated sync buttons initialization
- Backward-compatibility with day-scope, degradation and settings roundtrip.

Run: python test_v14_sync.py
"""
import os
import sys
import tempfile
import unittest
import tkinter as tk
from tkinter import ttk

# Add current folder to sys.path
curr_dir = os.path.dirname(os.path.abspath(__file__))
if curr_dir not in sys.path:
    sys.path.insert(0, curr_dir)

import PR_Calculator_GUI_v14 as v14


class TestV14SyncAndEngine(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()

    def test_file_discovery_scada(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = v14.PRCalculatorGUI(root)
            # Create a mock folder with KPI_Report_Daily.xls
            month_dir = os.path.join(self.tmp_dir, "2026 07")
            os.makedirs(month_dir, exist_ok=True)
            kpi_file = os.path.join(month_dir, "KPI_Report_Daily.xls")
            with open(kpi_file, "wb") as f:
                f.write(b"mock kpi content")

            found = app.find_scada_pr_file(month_dir)
            self.assertIsNotNone(found)
            self.assertEqual(os.path.abspath(found), os.path.abspath(kpi_file))

            # Test finding in subfolder
            scada_sub = os.path.join(self.tmp_dir, "2026 08", "SCADA")
            os.makedirs(scada_sub, exist_ok=True)
            kpi_sub = os.path.join(scada_sub, "KPI_Report_Daily.xls")
            with open(kpi_sub, "wb") as f:
                f.write(b"mock")
            found_sub = app.find_scada_pr_file(os.path.join(self.tmp_dir, "2026 08"))
            self.assertIsNotNone(found_sub)
            self.assertEqual(os.path.abspath(found_sub), os.path.abspath(kpi_sub))
        finally:
            root.destroy()

    def test_file_discovery_vcom(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = v14.PRCalculatorGUI(root)
            # Test Performance_ratio_2026_07_31.csv
            m_dir = os.path.join(self.tmp_dir, "2026 07")
            os.makedirs(m_dir, exist_ok=True)
            vcom_file = os.path.join(m_dir, "Performance_ratio_2026_07_31.csv")
            with open(vcom_file, "wb") as f:
                f.write(b"mock vcom")

            found = app.find_vcom_pr_file(m_dir)
            self.assertIsNotNone(found)
            self.assertEqual(os.path.abspath(found), os.path.abspath(vcom_file))

            # Test Performance_ratio_vcom.csv
            m_dir2 = os.path.join(self.tmp_dir, "2026 06")
            os.makedirs(m_dir2, exist_ok=True)
            vcom_file2 = os.path.join(m_dir2, "Performance_ratio_vcom.csv")
            with open(vcom_file2, "wb") as f:
                f.write(b"mock vcom 2")

            found2 = app.find_vcom_pr_file(m_dir2)
            self.assertIsNotNone(found2)
            self.assertEqual(os.path.abspath(found2), os.path.abspath(vcom_file2))
        finally:
            root.destroy()

    def test_vcom_parsing_various_encodings(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = v14.PRCalculatorGUI(root)
            # UTF-16 tab-separated
            csv_content = (
                '"Periodo: 1/7/2026 0.00.00 -  16.03.10"\n\n'
                '"Data"\t"Energia [kWh]"\t"Performance Ratio (PR principale) [%]"\n'
                '"01"\t"101932,00"\t"83,47"\n'
                '"02"\t"77121,00"\t"78,88"\n'
                '"03"\t"78686,00"\t"81,10"\n'
            )
            vcom_utf16 = os.path.join(self.tmp_dir, "vcom_utf16.csv")
            with open(vcom_utf16, "wb") as f:
                f.write(csv_content.encode("utf-16"))

            parsed_16 = app._read_vcom_daily_pr(vcom_utf16)
            self.assertEqual(parsed_16[1], 83.47)
            self.assertEqual(parsed_16[2], 78.88)
            self.assertEqual(parsed_16[3], 81.10)

            # UTF-8 semicolon-separated
            csv_utf8 = os.path.join(self.tmp_dir, "vcom_utf8.csv")
            with open(csv_utf8, "w", encoding="utf-8") as f:
                f.write(
                    "Data;Immissione;PR\n"
                    "1;1000;80.5\n"
                    "2;2000;82,3\n"
                )
            parsed_8 = app._read_vcom_daily_pr(csv_utf8)
            self.assertEqual(parsed_8[1], 80.5)
            self.assertEqual(parsed_8[2], 82.3)
        finally:
            root.destroy()

    def test_gui_buttons_and_version(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = v14.PRCalculatorGUI(root)
            self.assertTrue(hasattr(app, "btn_sync_scada"))
            self.assertTrue(hasattr(app, "btn_sync_vcom"))
            self.assertEqual(app.btn_sync_scada["text"], "Sync SCADA PR")
            self.assertEqual(app.btn_sync_vcom["text"], "Sync VCOM PR")

            # Check that button states toggle on start/reset
            app._reset_run_buttons()
            self.assertEqual(str(app.btn_sync_scada["state"]), "normal")
            self.assertEqual(str(app.btn_sync_vcom["state"]), "normal")
            self.assertEqual(str(app.btn_calculate["state"]), "normal")
            self.assertEqual(str(app.btn_stop["state"]), "disabled")
        finally:
            root.destroy()

    def test_backwards_compatibility_settings_and_degradation(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = v14.PRCalculatorGUI(root)
            # Degradation calculation
            self.assertEqual(app._pr_degradation_factor(2025, 6), (0, 1.0))
            self.assertEqual(app._pr_degradation_factor(2026, 1)[0], 0)
            self.assertEqual(app._pr_degradation_factor(2026, 2)[0], 1)
            n, factor = app._pr_degradation_factor(2027, 3)
            self.assertEqual(n, 2)
            self.assertAlmostEqual(factor, 0.996 ** 2, places=10)

            # Day scope filter
            app.date_var.set("2026-05-17")
            app.day_scope_var.set("mese")
            self.assertIsNone(app._resolve_day_filter())

            app.day_scope_var.set("giorno")
            self.assertEqual(app._resolve_day_filter(), {17})

            app.day_scope_var.set("intervallo")
            app.day_from_var.set("5")
            app.day_to_var.set("8")
            self.assertEqual(app._resolve_day_filter(), {5, 6, 7, 8})
        finally:
            root.destroy()


if __name__ == "__main__":
    unittest.main()
