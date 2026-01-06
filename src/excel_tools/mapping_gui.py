"""
mapping_gui.py

DEPRECATED: This file has been replaced by the advanced mapping functionality
integrated into import_dialog.py which uses the mapping_utils.py shared module.

The new implementation provides:
- Support for code/label-based mappings
- Formula building with operators (+, -, *, /)
- Integration with CustomMap sheet
- Better UI with improved styling

For new implementations, use:
    from src.excel_tools.import_dialog import launch_mapping_gui

This file is kept for backwards compatibility reference only.
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd


TARGET_LABELS = [
    "employee_id",
    "salaire_brut",
    "cot_salarie",
    "cot_patronale",
    "net_imposable",
    "PAS",
    "net_paye",
    "avantages",
    "total_cost",
]


def _read_source_sheets(path):
    # Read all sheets to get their DataFrames; return dict(sheet_name -> df)
    try:
        sheets = pd.read_excel(path, sheet_name=None)
        # normalize column names to strings
        for k, v in sheets.items():
            v.columns = v.columns.astype(str)
            sheets[k] = v
        return sheets
    except Exception:
        raise


def process_mapping(wb, source_path):
    """Launch mapping GUI and, on submit, create sheet in `wb` named after
    source file (without extension) and write mapped data.
    """
    try:
        sheets = _read_source_sheets(source_path)
        # pick the first sheet as preview for mapping UI (they have identical headers)
        first_sheet_name = next(iter(sheets.keys()))
        df = sheets[first_sheet_name]
        source_cols = list(df.columns.astype(str))
    except Exception as e:
        messagebox.showerror("Error reading file", f"Could not read {source_path}: {e}")
        return False

    root = tk.Tk()
    root.title("Mapping: map source columns to target fields")
    frm = ttk.Frame(root, padding=10)
    frm.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frm, text=f"Source: {os.path.basename(source_path)}").pack(anchor=tk.W)

    combo_vars = {}

    for label in TARGET_LABELS:
        row = ttk.Frame(frm)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text=label, width=18).pack(side=tk.LEFT)
        var = tk.StringVar()
        cb = ttk.Combobox(row, textvariable=var, values=["(none)"] + source_cols, width=50)
        cb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        combo_vars[label] = var

    def do_submit():
        mapping = {t: combo_vars[t].get() for t in TARGET_LABELS}
        if not mapping.get("employee_id") or mapping.get("employee_id") == "(none)":
            messagebox.showwarning("Missing mapping", "You must map 'employee_id' to a source column.")
            return

        # Build output DataFrame
        out = pd.DataFrame()
        emp_col = mapping["employee_id"]
        out["employee_id"] = df[emp_col]

        for t in TARGET_LABELS[1:]:
            src = mapping.get(t)
            if not src or src == "(none)":
                out[t] = ""
            else:
                # If the source column exists, copy it; otherwise blank
                if src in df.columns:
                    out[t] = df[src]
                else:
                    out[t] = ""

        # Create or replace sheets in workbook for every sheet in source
        try:
            for src_sheet_name, src_df in sheets.items():
                out = pd.DataFrame()
                emp_col = mapping["employee_id"]
                # if the chosen employee_id column is missing in this sheet, skip
                if emp_col not in src_df.columns:
                    messagebox.showwarning("Missing column", f"Sheet '{src_sheet_name}' does not contain column '{emp_col}'. Skipping this sheet.")
                    continue
                out["employee_id"] = src_df[emp_col]

                for t in TARGET_LABELS[1:]:
                    src = mapping.get(t)
                    if not src or src == "(none)":
                        out[t] = ""
                    else:
                        if src in src_df.columns:
                            out[t] = src_df[src]
                        else:
                            out[t] = ""

                # sheet name in target = source sheet name (truncate to 31)
                sheet_name = str(src_sheet_name)[:31]
                try:
                    # remove if exists
                    try:
                        existing = wb.sheets[sheet_name]
                        existing.delete()
                    except Exception:
                        pass

                    sh = wb.sheets.add(sheet_name)
                    # write header with formatting (mimic screenshot: bold, yellow background)
                    sh.range((1, 1)).value = list(out.columns)
                    header_range = sh.range((1, 1), (1, len(out.columns)))
                    header_range.api.Font.Bold = True
                    header_range.color = (255, 204, 0)  # light yellow
                    # write data starting row 2
                    sh.range((2, 1)).options(index=False, header=False).value = out

                    # Autosize columns
                    for i in range(1, len(out.columns) + 1):
                        sh.range((1, i)).api.EntireColumn.AutoFit()
                except Exception as e:
                    messagebox.showerror("Error writing sheet", f"Could not write sheet '{sheet_name}': {e}")
                    # continue with next sheet
                    continue

            wb.save()
            messagebox.showinfo("Done", f"Sheets created/updated in Control Panel workbook from '{os.path.basename(source_path)}'.")
        except Exception as e:
            messagebox.showerror("Error writing sheets", f"An error occurred while writing sheets: {e}")
        finally:
            root.destroy()

    btn_frame = ttk.Frame(frm)
    btn_frame.pack(fill=tk.X, pady=6)
    ttk.Button(btn_frame, text="Submit", command=do_submit).pack(side=tk.RIGHT, padx=6)
    ttk.Button(btn_frame, text="Cancel", command=root.destroy).pack(side=tk.RIGHT)

    root.mainloop()
    return True
