import os
import sys
import tempfile
import shutil
import unittest
import pandas as pd

sys.path.insert(0, r"\\s01\get\2025.01 Mazara 01 A2A\03 - REPORT\Report\09 Testing\PR Calculation automation")
import PR_Calculator_GUI_v14 as v14
from VCOM_to_SCADA import convert_vcom_to_scada

class TestVCOMAndMixedBatch(unittest.TestCase):

    def setUp(self):
        self.vcom_test_data = r"\\S01\get\2025.01 Mazara 01 A2A\03 - REPORT\Report\09 Testing\PR Calculation automation\test_data_vcom"
        self.scada_15_dir = r"\\s01\get\2025.01 Mazara 01 A2A\03 - REPORT\Report\01 Daily Reports\2026 08\15"

    def test_01_vcom_to_scada_all_7_files(self):
        """Test that VCOM_to_SCADA generates all 7 required pseudo-SCADA files."""
        out_dir = tempfile.mkdtemp()
        try:
            convert_vcom_to_scada(self.vcom_test_data, out_dir, "2026-08-15")
            files = sorted(os.listdir(out_dir))
            expected = [
                "Regolazione_della_potenza_attiva_2026_08_15.xlsx",
                "SATAC_Meter_15Min.xlsx",
                "TS_01_Inverter_15Min.xlsx",
                "TS_01_Weather_15Min.xlsx",
                "TS_02_Inverter_15Min.xlsx",
                "TS_03_Inverter_15Min.xlsx",
                "TS_03_Weather_15Min.xlsx"
            ]
            self.assertEqual(len(files), 7)
            for exp in expected:
                self.assertIn(exp, files)
        finally:
            shutil.rmtree(out_dir, ignore_errors=True)

    def test_02_vcom_days_parsing(self):
        """Test parsing of VCOM day string into integer sets."""
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        app = v14.PRCalculatorGUI(root)
        
        app.vcom_days_var.set("3, 7, 9")
        self.assertEqual(app._parse_vcom_days_set(), {3, 7, 9})
        
        app.vcom_days_var.set("1-4, 8, 15")
        self.assertEqual(app._parse_vcom_days_set(), {1, 2, 3, 4, 8, 15})
        
        app.vcom_days_var.set(" 5 ; 10 ; 20 ")
        self.assertEqual(app._parse_vcom_days_set(), {5, 10, 20})
        
        app.vcom_days_var.set("")
        self.assertEqual(app._parse_vcom_days_set(), set())
        
        root.destroy()

    def test_03_single_day_vcom_calculation(self):
        """Test calculate_single_day on VCOM pseudo-SCADA dataset."""
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        app = v14.PRCalculatorGUI(root)
        app.backup_mother_var.set(False)
        app.sync_mother_var.set(False)

        out_dir = tempfile.mkdtemp()
        try:
            convert_vcom_to_scada(self.vcom_test_data, out_dir, "2026-08-15")
            df_res, calc_res = app.calculate_single_day(
                out_dir, "2026-08-15", 0.828, 50, diff_threshold=0.10,
                calcolo_folder=out_dir, skip_mother_update=True, poa_method="average"
            )
            self.assertIsNotNone(df_res)
            self.assertIsNotNone(calc_res)
            # Verify PR value is near the benchmark 86.29%
            comp_pr = calc_res.get('comp_raw_pr', 0.0)
            self.assertGreater(comp_pr, 80.0)
            self.assertLess(comp_pr, 95.0)
            self.assertAlmostEqual(comp_pr, 86.287, delta=0.5)
        finally:
            shutil.rmtree(out_dir, ignore_errors=True)
            root.destroy()

    def test_04_mixed_month_batch_simulation(self):
        """Test mixed month execution where Day 1 uses SCADA and Day 15 uses VCOM."""
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        app = v14.PRCalculatorGUI(root)
        app.backup_mother_var.set(False)
        app.sync_mother_var.set(False)

        sim_month_dir = tempfile.mkdtemp()
        try:
            # Create Day 01 (SCADA folder)
            d01_dir = os.path.join(sim_month_dir, "01")
            shutil.copytree(self.scada_15_dir, d01_dir)
            # rename 2026_08_15 in reg to 2026_08_01
            for f in os.listdir(d01_dir):
                if "2026_08_15" in f:
                    os.rename(os.path.join(d01_dir, f), os.path.join(d01_dir, f.replace("2026_08_15", "2026_08_01")))

            # Create Day 15 (VCOM folder with vcom subfolder)
            d15_dir = os.path.join(sim_month_dir, "15")
            os.makedirs(os.path.join(d15_dir, "vcom"), exist_ok=True)
            for f in os.listdir(self.vcom_test_data):
                if f.endswith(".csv"):
                    shutil.copy(os.path.join(self.vcom_test_data, f), os.path.join(d15_dir, "vcom", f))

            # Configure App for Mixed mode with day 15 as VCOM
            app.folder_path_var.set(sim_month_dir)
            app.data_source_var.set("misto")
            app.vcom_days_var.set("15")
            app.day_scope_var.set("mese")

            # Check that day 15 is detected as VCOM
            vcom_days = app._parse_vcom_days_set()
            self.assertIn(15, vcom_days)

            # Test VCOM folder detection for Day 15
            found_vcom = app._find_vcom_folder_for_day(d15_dir)
            self.assertIsNotNone(found_vcom)
            self.assertTrue(os.path.exists(found_vcom))

        finally:
            shutil.rmtree(sim_month_dir, ignore_errors=True)
            root.destroy()

    def test_05_vcom_mother_row_orange_highlight(self):
        """Test that days calculated via VCOM are highlighted in Light Orange in the Mother file."""
        import tkinter as tk
        import win32com.client

        root = tk.Tk()
        root.withdraw()
        app = v14.PRCalculatorGUI(root)
        app.backup_mother_var.set(False)
        app.sync_mother_var.set(True)

        sim_dir = tempfile.mkdtemp()
        try:
            calcolo_folder = os.path.join(sim_dir, "PR CALCOLO FILE")
            os.makedirs(calcolo_folder, exist_ok=True)

            template_src = os.path.join(r"\\s01\get\2025.01 Mazara 01 A2A\03 - REPORT\Report\09 Testing\PR Calculation automation\original_format", "PR_recalculation_26_apr.xlsx")
            mother_src = os.path.join(r"\\s01\get\2025.01 Mazara 01 A2A\03 - REPORT\Report\09 Testing\PR Calculation automation\original_format", "00 PR_recalculation_APRL.xlsx")

            shutil.copy2(template_src, os.path.join(calcolo_folder, "PR_recalculation_04_ago.xlsx"))
            shutil.copy2(template_src, os.path.join(calcolo_folder, "PR_recalculation_05_ago.xlsx"))
            shutil.copy2(mother_src, os.path.join(calcolo_folder, "00 PR_recalculation_AGOS.xlsx"))

            # Sync with Day 5 as VCOM
            app.sync_mother_file(calcolo_folder, 2026, 8, vcom_days={5})

            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(os.path.join(calcolo_folder, "00 PR_recalculation_AGOS.xlsx"))
            ws = wb.Sheets(1)

            color_day4 = ws.Cells(8, 1).Interior.Color
            color_day5 = ws.Cells(9, 1).Interior.Color
            wb.Close(SaveChanges=False)
            excel.Quit()

            # Day 5 (VCOM) must be light orange (11722975), Day 4 (SCADA) must not be orange
            self.assertEqual(int(color_day5), 11722975)
            self.assertNotEqual(int(color_day4), 11722975)

        finally:
            shutil.rmtree(sim_dir, ignore_errors=True)
            root.destroy()


if __name__ == "__main__":
    unittest.main()
