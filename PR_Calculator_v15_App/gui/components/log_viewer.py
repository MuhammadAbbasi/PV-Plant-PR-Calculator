import sys
import tkinter as tk
from tkinter import ttk

class LogRedirector:
    """Redirects stdout and stderr streams cleanly to a Tkinter Text widget."""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, str_val):
        try:
            self.text_widget.insert(tk.END, str_val)
            self.text_widget.see(tk.END)
        except Exception:
            pass

    def flush(self):
        pass
