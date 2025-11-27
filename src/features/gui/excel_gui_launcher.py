"""
excel_gui_launcher.py
Excel entry point with GUI for defining custom mappings.

Pipeline:
1. Run fill_simplified_table() to populate Sheet1/Sheet2.
2. Apply any previously saved mappings from CustomMap.
3. Open a Tkinter GUI so the user can add / edit a mapping.
"""

import xlwings as xw
import tkinter as tk
from tkinter import ttk, messagebox

from src.features.payroll_processing.automation import fill_simplified_table
from src.features.payroll_processing.mapping_engine import (
    apply_saved_mappings,
    apply_single_mapping,
    reset_to_default,
    save_or_update_mapping_row,
)
from src.core.normalization import _norm_label
from src.core.config import TGT_SHEET
from src.core.data_loader import get_feuil1_rows_info


def _get_workbook() -> xw.Book:
    try:
        wb = xw.Book.caller()
    except Exception:
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        wb = app.books[0]
    return wb


def show_mapping_gui(wb: xw.Book):
    sh_tgt = wb.sheets[TGT_SHEET]

    headers = sh_tgt.range("A1").expand("right").value
    if not isinstance(headers, list):
        headers = [headers]
    sheet1_labels = headers[1:]  # skip employee ID column

    rows_info, code_to_rows, _ = get_feuil1_rows_info(wb)
    codes_sorted = sorted(code_to_rows.keys())

    labels_display = []
    map_display_to_raw = {}
    for r in rows_info:
        raw = r["label_raw"]
        code = r["code"]
        if code:
            disp = f"{raw} ({code})"
        else:
            disp = raw
        labels_display.append(disp)
        map_display_to_raw[disp] = raw

    labels_display = sorted(set(labels_display))

    root = tk.Tk()
    root.title("Custom mapping")

    frame = ttk.Frame(root, padding=10)
    frame.grid(sticky="nsew")

    ttk.Label(frame, text="1. Select a label from Sheet1:").grid(row=0, sticky="w")

    var_label = tk.StringVar()
    cb_sheet1 = ttk.Combobox(frame, textvariable=var_label, state="readonly")
    cb_sheet1["values"] = sheet1_labels
    cb_sheet1.grid(row=1, columnspan=3, sticky="ew", pady=(0, 10))

    ttk.Label(frame, text="2. Operator for next item:").grid(row=2, sticky="w")
    var_op = tk.StringVar(value="+")
    cb_op = ttk.Combobox(frame, textvariable=var_op, state="readonly", width=5)
    cb_op["values"] = ["+", "-", "*", "/"]
    cb_op.grid(row=2, column=1, sticky="w")

    items = []  # unified list: (key, op, type)

    lbl_selected = ttk.Label(frame, text="Selected items: (none)")
    lbl_selected.grid(row=8, columnspan=3, sticky="w")

    # ----- codes -----
    ttk.Label(frame, text="3a. Add code:").grid(row=3, sticky="w", pady=(10, 0))
    var_code = tk.StringVar()
    cb_code = ttk.Combobox(frame, textvariable=var_code, state="readonly")
    cb_code["values"] = codes_sorted
    cb_code.grid(row=4, sticky="ew")

    lbl_code_feedback = ttk.Label(frame, text="", foreground="green")
    lbl_code_feedback.grid(row=5, columnspan=3, sticky="w")

    def add_code():
        c = var_code.get()
        if not c:
            return
        op = var_op.get()
        entry = (c, op, "code")
        if entry not in items:
            items.append(entry)
            lbl_selected["text"] = "Selected: " + ", ".join(
                f"{op}{key}" for key, op, _ in items
            )
            lbl_code_feedback["text"] = "✔ Code added"
            lbl_code_feedback.after(1200, lambda: lbl_code_feedback.config(text=""))

    ttk.Button(frame, text="Add code", command=add_code).grid(row=4, column=1, padx=5)

    # ----- labels -----
    ttk.Label(frame, text="3b. Add label:").grid(row=6, sticky="w", pady=(10, 0))
    var_lbl = tk.StringVar()
    cb_lbl = ttk.Combobox(frame, textvariable=var_lbl, state="readonly")
    cb_lbl["values"] = labels_display
    cb_lbl.grid(row=7, sticky="ew")

    lbl_lbl_feedback = ttk.Label(frame, text="", foreground="green")
    lbl_lbl_feedback.grid(row=9, columnspan=3, sticky="w")

    def add_label():
        disp = var_lbl.get()
        if not disp:
            return
        raw = map_display_to_raw.get(disp, disp)
        op = var_op.get()
        entry = (raw, op, "label")
        if entry not in items:
            items.append(entry)
            lbl_selected["text"] = "Selected: " + ", ".join(
                f"{op}{key}" for key, op, _ in items
            )
            lbl_lbl_feedback["text"] = "✔ Label added"
            lbl_lbl_feedback.after(1200, lambda: lbl_lbl_feedback.config(text=""))

    ttk.Button(frame, text="Add label", command=add_label).grid(row=7, column=1, padx=5)

    ttk.Label(frame, text="4. Save this mapping?").grid(row=10, sticky="w", pady=(10, 0))
    var_save = tk.BooleanVar(value=True)
    ttk.Checkbutton(
        frame,
        text="Save for next time",
        variable=var_save,
    ).grid(row=11, sticky="w")

    def reset_defaults():
        if not messagebox.askyesno("Reset", "Reset to default?"):
            return

        reset_to_default(wb)
        fill_simplified_table()
        wb.save()
        messagebox.showinfo("Reset", "Default restored.")
        root.destroy()

    def apply_mapping():
        sheet1_lbl = var_label.get()
        if not sheet1_lbl:
            messagebox.showerror("Error", "Select a Sheet1 label.")
            return

        if not items:
            messagebox.showerror("Error", "Add at least one item.")
            return

        apply_single_mapping(wb, sheet1_label_raw=sheet1_lbl, keys_ops=items)

        if var_save.get():
            save_or_update_mapping_row(wb, sheet1_lbl, items)
            wb.save()
            messagebox.showinfo("OK", "Saved and applied.")
        else:
            messagebox.showinfo("OK", "Applied for this run only.")

        root.destroy()

    ttk.Button(frame, text="Reset to default", command=reset_defaults).grid(
        row=12, column=0, pady=10
    )
    ttk.Button(frame, text="Apply", command=apply_mapping).grid(
        row=12, column=1, pady=10
    )
    ttk.Button(frame, text="Cancel", command=root.destroy).grid(
        row=12, column=2, pady=10
    )

    root.mainloop()


def main():
    wb = _get_workbook()
    fill_simplified_table()
    apply_saved_mappings(wb)
    show_mapping_gui(wb)


if __name__ == "__main__":
    main()
