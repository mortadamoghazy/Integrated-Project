import tkinter as tk
from tkinter import ttk, messagebox
import xlwings as xw

from src.core.normalization import _norm_label
from src.core.config import TGT_SHEET
from src.features.analytics.plot_salary_field import plot_field


def _get_workbook():
    try:
        wb = xw.Book.caller()
    except Exception:
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        wb = app.books[0]
    return wb


def main():
    wb = _get_workbook()
    sh = wb.sheets[TGT_SHEET]

    # Load column headers from Sheet1 (fields)
    headers = sh.range("A1").expand("right").value
    headers = headers[1:]  # skip employee ID column

    root = tk.Tk()
    root.title("Plot Field Values")

    frame = ttk.Frame(root, padding=10)
    frame.grid(sticky="nsew")

    ttk.Label(frame, text="Select a field to plot:").grid(row=0, sticky="w")

    var_label = tk.StringVar()
    cb = ttk.Combobox(frame, textvariable=var_label, state="readonly")
    cb["values"] = headers
    cb.grid(row=1, sticky="ew", pady=10)

    def do_plot():
        lbl = var_label.get()
        if not lbl:
            messagebox.showerror("Error", "Select a field")
            return
        try:
            plot_field(wb, lbl)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    ttk.Button(frame, text="Plot", command=do_plot).grid(row=2, sticky="ew")
    ttk.Button(frame, text="Close", command=root.destroy).grid(row=3, sticky="ew")

    root.mainloop()
