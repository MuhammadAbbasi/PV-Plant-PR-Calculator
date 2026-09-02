import sys
import os
import ctypes

# Ensure parent root directory is in Python search path
pkg_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(pkg_dir)

if project_root not in sys.path:
    sys.path.insert(0, project_root)
if pkg_dir not in sys.path:
    sys.path.insert(0, pkg_dir)

import tkinter as tk

# Enable High DPI Awareness on Windows
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

from PR_Calculator_GUI_v15 import PRCalculatorGUI

def main():
    root = tk.Tk()
    app = PRCalculatorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
