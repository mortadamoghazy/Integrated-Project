"""
tests/test_monthly_mapping_gui_launcher.py

GUI for defining global mappings (once) to be reused across all monthly sheets.
- No automation. No default mapping.
- Destination labels are taken from the template workbook.
- Source row codes/labels are taken from a user-selected monthly source sheet.
- Mappings are stored in CustomMap in the source workbook.
"""

from __future__ import annotations

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Tuple, Dict

import xlwings as xw

from src.excel_tools.mapping_utils import (
    ensure_custom_sheet,
    serialize_keys_ops,
    get_source_rows_display,
    save_mapping_row,
    get_month_sheets,
    is_month_sheet,
    CUSTOM_SHEET,
    DEFAULT_MONTH_SHEET_REGEX
)
from tests.test_monthly_mapping_engine import (
    SOURCE_EXCEL_PATH,
    DEST_HEADERS,
)


def open_template_and_get_headers(app: xw.App, template_path: str, template_sheet: str) -> List[str]:
    """Open template workbook and extract headers (kept for backwards compatibility)."""
    template_wb = None
    for b in app.books:
        if b.fullname.lower() == template_path.lower():
            template_wb = b
            break
    if template_wb is None:
        template_wb = app.books.open(template_path)

    if template_sheet not in [s.name for s in template_wb.sheets]:
        raise ValueError(f"Template sheet '{template_sheet}' not found in template workbook.")

    sh = template_wb.sheets[template_sheet]
    headers = sh.range("A1").expand("right").value
    if not isinstance(headers, list):
        headers = [headers]

    headers = [h for h in headers if h not in (None, "")]
    return headers


def main(
    month_regex: str = DEFAULT_MONTH_SHEET_REGEX,
    label_col: int = 2,
):
    # Always use the source workbook so mappings persist across runs
    try:
        wb = xw.Book.caller()
        # If caller exists but is not the source, open source explicitly
        if wb.name.lower() != os.path.basename(SOURCE_EXCEL_PATH).lower():
            app = wb.app
            # Try reuse an already-open source if present
            src = None
            for b in app.books:
                if b.name.lower() == os.path.basename(SOURCE_EXCEL_PATH).lower():
                    src = b
                    break
            wb = src if src is not None else app.books.open(SOURCE_EXCEL_PATH)
    except Exception:
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        # Prefer an already-open source workbook
        src = None
        for b in app.books:
            try:
                if b.name.lower() == os.path.basename(SOURCE_EXCEL_PATH).lower():
                    src = b
                    break
            except Exception:
                pass
        wb = src if src is not None else app.books.open(SOURCE_EXCEL_PATH)

    ensure_custom_sheet(wb)

    # Use DEST_HEADERS as template headers, excluding employee_id (auto-mapped)
    template_headers = [h for h in DEST_HEADERS if h.lower() != "employee_id"]

    month_sheets = get_month_sheets(wb, month_regex=month_regex)
    if not month_sheets:
        raise RuntimeError(f"No month sheets found matching regex: {month_regex}")

    # GUI with improved styling
    root = tk.Tk()
    root.title("Payroll Mapping Configuration")
    root.geometry("800x650")
    root.resizable(True, True)
    root.configure(bg="#f0f0f0")
    
    # Configure style
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('Title.TLabel', font=('Segoe UI', 11, 'bold'), foreground='#2c3e50')
    style.configure('TLabel', font=('Segoe UI', 9), foreground='#34495e')
    style.configure('TCombobox', font=('Segoe UI', 9))
    style.configure('TButton', font=('Segoe UI', 9), padding=6)
    
    # Main frame with better styling
    main_frame = ttk.Frame(root, padding=20)
    main_frame.grid(sticky="nsew", padx=10, pady=10)
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    
    # Header
    header = ttk.Label(main_frame, text="Payroll Data Mapping Configuration", 
                      style='Title.TLabel')
    header.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky="w")
    
    # Separator
    ttk.Separator(main_frame, orient='horizontal').grid(row=1, column=0, columnspan=2, 
                                                         sticky="ew", pady=(0, 15))
    
    frame = main_frame

    ttk.Label(frame, text="1) Source Month Sheet:", font=('Segoe UI', 9, 'bold')).grid(row=2, column=0, sticky="w", pady=(5,2))
    var_src_sheet = tk.StringVar(value=month_sheets[0])
    cb_src_sheet = ttk.Combobox(frame, textvariable=var_src_sheet, state="readonly", values=month_sheets, width=30)
    cb_src_sheet.grid(row=3, column=0, sticky="ew", pady=(0, 15))

    ttk.Label(frame, text="2) Destination Column:", font=('Segoe UI', 9, 'bold')).grid(row=4, column=0, sticky="w", pady=(5,2))
    var_dest = tk.StringVar()
    cb_dest = ttk.Combobox(frame, textvariable=var_dest, state="readonly", values=template_headers, width=50)
    cb_dest.grid(row=5, column=0, sticky="ew", pady=(0, 15))

    ttk.Label(frame, text="3) Operator:", font=('Segoe UI', 9, 'bold')).grid(row=6, column=0, sticky="w", pady=(5,2))
    var_op = tk.StringVar(value="+")
    cb_op = ttk.Combobox(frame, textvariable=var_op, state="readonly", values=["+", "-", "*", "/"], width=10)
    cb_op.grid(row=7, column=0, sticky="w", pady=(0, 15))

    ttk.Label(frame, text="4) Build Formula - Add by Code:", font=('Segoe UI', 9, 'bold')).grid(row=8, column=0, sticky="w", pady=(5,2))
    var_code = tk.StringVar()
    cb_code = ttk.Combobox(frame, textvariable=var_code, state="readonly", width=35)
    cb_code.grid(row=9, column=0, sticky="ew")

    ttk.Label(frame, text="Or Add by Label:", font=('Segoe UI', 9)).grid(row=10, column=0, sticky="w", pady=(10,2))
    var_label = tk.StringVar()
    cb_label = ttk.Combobox(frame, textvariable=var_label, state="readonly", width=70)
    cb_label.grid(row=11, column=0, sticky="ew")

    items: List[Tuple[str, str, str]] = []  # (key, op, type)
    
    # Current formula display with background
    formula_frame = tk.Frame(frame, bg='#ecf0f1', relief='solid', borderwidth=1)
    formula_frame.grid(row=12, column=0, sticky="ew", pady=(15, 10), padx=2)
    ttk.Label(formula_frame, text="Current Formula:", font=('Segoe UI', 8, 'bold'), 
             background='#ecf0f1').pack(anchor='w', padx=8, pady=(5,2))
    lbl_selected = ttk.Label(formula_frame, text="(none)", font=('Consolas', 9), 
                            background='#ecf0f1', foreground='#27ae60')
    lbl_selected.pack(anchor='w', padx=8, pady=(0, 5))

    display_to_raw: Dict[str, str] = {}

    def refresh_source_lists():
        sh = wb.sheets[var_src_sheet.get()]
        labels_display, map_disp_to_raw, codes_sorted = get_source_rows_display(sh, label_col=label_col)
        nonlocal display_to_raw
        display_to_raw = map_disp_to_raw
        cb_code["values"] = codes_sorted
        cb_label["values"] = labels_display

    def add_code():
        c = var_code.get()
        if not c:
            return
        op = var_op.get()
        entry = (c, op, "code")
        if entry not in items:
            items.append(entry)
        lbl_selected["text"] = "Selected items: " + ", ".join([f"{op}{key}" for key, op, _ in items])

    def add_label():
        disp = var_label.get()
        if not disp:
            return
        raw = display_to_raw.get(disp, disp)
        op = var_op.get()
        entry = (raw, op, "label")
        if entry not in items:
            items.append(entry)
        lbl_selected["text"] = "Selected items: " + ", ".join([f"{op}{key}" for key, op, _ in items])

    def clear_items():
        items.clear()
        lbl_selected["text"] = "Selected items: (none)"

    def save_rule():
        dest = var_dest.get()
        if not dest:
            messagebox.showerror("Error", "Select a destination column from template.")
            return
        if not items:
            messagebox.showerror("Error", "Add at least one code or label item.")
            return

        save_mapping_row(wb, dest_label_raw=dest, keys_ops=items)
        wb.save()
        messagebox.showinfo("Saved", f"Saved mapping for '{dest}' in {CUSTOM_SHEET} sheet.")
        clear_items()

    # Button styling
    style.configure('Action.TButton', font=('Segoe UI', 9), padding=8)
    style.configure('Primary.TButton', font=('Segoe UI', 9, 'bold'), padding=8)
    
    # Buttons layout
    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=13, column=0, sticky="ew", pady=(15, 0))
    
    ttk.Button(btn_frame, text="🔄 Refresh Lists", command=refresh_source_lists, 
              style='Action.TButton').pack(side='left', padx=(0, 5))
    ttk.Button(btn_frame, text="➕ Add Code", command=add_code, 
              style='Action.TButton').pack(side='left', padx=5)
    ttk.Button(btn_frame, text="➕ Add Label", command=add_label, 
              style='Action.TButton').pack(side='left', padx=5)
    ttk.Button(btn_frame, text="🗑 Clear", command=clear_items, 
              style='Action.TButton').pack(side='left', padx=5)
    
    # Bottom action buttons
    bottom_frame = ttk.Frame(frame)
    bottom_frame.grid(row=14, column=0, sticky="ew", pady=(20, 0))
    
    ttk.Button(bottom_frame, text="💾 Save Mapping Rule", command=save_rule, 
              style='Primary.TButton').pack(side='left', padx=(0, 10))
    ttk.Button(bottom_frame, text="❌ Close", command=root.destroy, 
              style='Action.TButton').pack(side='right')

    refresh_source_lists()
    root.mainloop()


if __name__ == "__main__":
    main()
