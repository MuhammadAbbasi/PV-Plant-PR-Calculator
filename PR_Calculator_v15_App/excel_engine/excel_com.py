import win32com.client
import os

def get_excel_app():
    """Dispatches or retrieves a silent Excel COM application."""
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    return excel

def open_workbook_writable(excel, file_path):
    """Opens workbook with read-only write access safely."""
    abs_path = os.path.abspath(file_path).replace('/', '\\')
    wb = excel.Workbooks.Open(
        Filename=abs_path, UpdateLinks=0, ReadOnly=False, Format=None,
        Password="", WriteResPassword="", IgnoreReadOnlyRecommended=True,
        Origin=None, Delimiter=None, Editable=True, Notify=False, Converter=None, AddToMru=False
    )
    return wb
