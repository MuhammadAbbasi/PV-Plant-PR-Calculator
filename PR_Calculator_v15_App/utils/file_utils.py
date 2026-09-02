import os
import glob
import re

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def find_file_by_patterns(folder, patterns):
    """Find first file matching any of the pattern strings inside a folder."""
    if not os.path.exists(folder):
        return None
    for pat in patterns:
        matches = glob.glob(os.path.join(folder, pat))
        if matches:
            return matches[0]
    return None
