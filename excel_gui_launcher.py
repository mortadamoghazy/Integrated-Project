"""
excel_gui_launcher.py
---------------------

Entry point for the Excel button.

When the button is pressed:
1. Runs the existing automation fill_simplified_table() from excel_automation_script.py
2. Applies any previously saved custom mappings (sheet 'CustomMap')
3. Opens a GUI so the user can define a new mapping:
   - Select a label in Sheet1
   - Add any combination of:
        * Codes (column A of Feuil1)
        * Labels (column B of Feuil1)
   - Each with operator: +, -, *, /
   - Save mapping for future runs or use once
   - Reset default fill (delete all mappings)
"""

import xlwings as xw
import tkinter as tk
from tkinter import ttk, messagebox

from excel_automation_script import fill_simplified_table, _norm_label, _norm_emp_id

SRC_SHEET = "Feuil1"
TGT_SHEET = "Sheet1"
CUSTOM_SHEET = "CustomMap"

BLOCK_WIDTH = 3
MAX_COLS = 120


# ----------------- Workbook helpers ----------------- #

def _get_workbook():
    try:
        wb = xw.Book.caller()
    except Exception:
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        wb = app.books[0]
    return wb


def _ensure_custom_sheet(wb):
    try:
        sh = wb.sheets[CUSTOM_SHEET]
    except Exception:
        sh = wb.sheets.add(CUSTOM_SHEET)
        sh.range("A1").value = ["Sheet1_Label", "Mapping_Type", "Feuil1_Keys"]
    return sh


# ----------------- Serialization (now with type) ----------------- #

def serialize_keys_ops(items):
    """
    items = [(key, op, type)]
    type ∈ {"code", "label"}
    Stored format:
       "+::F07::code;;-::Salaire de base::label"
    """
    return ";;".join(f"{op}::{key}::{item_type}" for key, op, item_type in items)


def deserialize_keys_ops(cell_value):
    """
    Returns list of (key, op, type)
    """
    result = []
    if not cell_value:
        return result

    parts = str(cell_value).split(";;")
    for entry in parts:
        if "::" not in entry:
            continue
        try:
            op, key, item_type = entry.split("::")
        except:
            continue
        op = op.strip()
        key = key.strip()
        item_type = item_type.strip().lower()
        if op in ["+", "-", "*", "/"] and key and item_type in ["code", "label"]:
            result.append((key, op, item_type))
    return result


# ----------------- CustomMap I/O ----------------- #

def load_saved_mappings(wb):
    try:
        sh = wb.sheets[CUSTOM_SHEET]
    except:
        return []

    used = sh.range("A1").current_region
    if used.rows.count <= 1:
        return []

    rows = used.value
    rules = []
    for row in rows[1:]:
        if not row or not row[0]:
            continue

        sheet1_label = row[0]
        mapping_type = row[1] or ""  # unused but kept for compatibility
        keys_cell = row[2]

        keys_ops = deserialize_keys_ops(keys_cell)
        if not keys_ops:
            continue

        rules.append({
            "sheet1_label_raw": sheet1_label,
            "sheet1_label_norm": _norm_label(sheet1_label),
            "keys_ops": keys_ops
        })

    return rules


def save_or_update_mapping_row(wb, sheet1_label_raw, keys_ops):
    sh = _ensure_custom_sheet(wb)
    used = sh.range("A1").current_region
    rows = used.rows.count

    target_norm = _norm_label(sheet1_label_raw)
    existing = None

    if rows > 1:
        data = used.value
        for idx, row in enumerate(data[1:], start=2):
            if row and row[0] and _norm_label(row[0]) == target_norm:
                existing = idx
                break

    if existing is None:
        existing = rows + 1

    key_string = serialize_keys_ops(keys_ops)

    sh.range(existing, 1).value = sheet1_label_raw
    sh.range(existing, 2).value = "mixed"
    sh.range(existing, 3).value = key_string


# ----------------- Feuil1 scanning ----------------- #

def read_feuil1_meta(wb):
    sh = wb.sheets[SRC_SHEET]

    last_row = max(
        sh.range("A" + str(sh.cells.rows.count)).end("up").row,
        sh.range("B" + str(sh.cells.rows.count)).end("up").row
    )

    rows_info = []
    code_to_rows = {}
    labelnorm_to_rows = {}

    for r in range(1, last_row + 1):
        code = sh.range((r, 1)).value
        label_raw = sh.range((r, 2)).value

        if not code and not label_raw:
            continue

        code_str = str(code).strip() if code else ""
        label_str = str(label_raw).strip() if label_raw else ""
        label_norm = _norm_label(label_str)

        rows_info.append({
            "row_index": r,
            "code": code_str,
            "label_raw": label_str,
            "label_norm": label_norm
        })

        if code_str:
            code_to_rows.setdefault(code_str, []).append(r)
        if label_norm:
            labelnorm_to_rows.setdefault(label_norm, []).append(r)

    return rows_info, code_to_rows, labelnorm_to_rows


# ----------------- Employee mapping ----------------- #

def read_employee_mappings(wb):
    sh_src = wb.sheets[SRC_SHEET]
    sh_tgt = wb.sheets[TGT_SHEET]

    headers = sh_tgt.range("A1").expand("right").value
    if not isinstance(headers, list):
        headers = [headers]
    header_map = {_norm_label(h): idx + 1 for idx, h in enumerate(headers)}

    ids = sh_tgt.range("A2").expand("down").value
    if not isinstance(ids, list):
        ids = [ids]
    ids = [str(v).strip() if v else "" for v in ids]

    digit_lengths = [len("".join(filter(str.isdigit, v))) for v in ids if v]
    id_width = max(digit_lengths) if digit_lengths else 5

    ids_norm = [_norm_emp_id(v, id_width) for v in ids]

    emp_to_sheet1 = {emp: i for i, emp in enumerate(ids_norm, start=2) if emp}

    row_vals = sh_src.range((3, 1), (3, MAX_COLS)).value
    if not isinstance(row_vals, list):
        row_vals = [row_vals]

    emp_to_feuil1 = {}
    for col, emp in enumerate(row_vals, start=1):
        if not emp:
            continue
        emp_norm = _norm_emp_id(emp, id_width)
        if emp_norm:
            emp_to_feuil1[emp_norm] = col

    return emp_to_feuil1, emp_to_sheet1, header_map


# ----------------- Summation helper ----------------- #

def _sum_rows_for_employee(sh_src, rows, col_start):
    total = 0
    for r in rows:
        vals = sh_src.range((r, col_start), (r, col_start + BLOCK_WIDTH - 1)).value
        if isinstance(vals, list):
            total += sum(v for v in vals if isinstance(v, (int, float)))
        else:
            if isinstance(vals, (int, float)):
                total += vals
    return total


# ----------------- Apply mixed mapping ----------------- #

def apply_single_mapping(wb, sheet1_label_raw, keys_ops):
    sh_src = wb.sheets[SRC_SHEET]
    sh_tgt = wb.sheets[TGT_SHEET]

    _, code_to_rows, labelnorm_to_rows = read_feuil1_meta(wb)
    emp_to_f1, emp_to_s1, header_map = read_employee_mappings(wb)

    target_col = header_map.get(_norm_label(sheet1_label_raw))
    if not target_col:
        return

    for emp_norm, col_start in emp_to_f1.items():
        row_s1 = emp_to_s1.get(emp_norm)
        if not row_s1:
            continue

        total = None

        for key, op, item_type in keys_ops:
            if item_type == "code":
                rows = code_to_rows.get(key, [])
            else:
                rows = labelnorm_to_rows.get(_norm_label(key), [])

            if not rows:
                continue

            part = _sum_rows_for_employee(sh_src, rows, col_start)

            if total is None:
                total = 0 if op in ["+", "-"] else 1

            if op == "+": total += part
            elif op == "-": total -= part
            elif op == "*": total *= part
            elif op == "/" and part != 0: total /= part

        if total is not None:
            cell = sh_tgt.range((row_s1, target_col))
            cell.value = total
            cell.color = (204, 255, 204)
            cell.api.Font.Color = 0
            cell.api.Font.Bold = True


def apply_saved_mappings(wb):
    for rule in load_saved_mappings(wb):
        apply_single_mapping(
            wb,
            sheet1_label_raw=rule["sheet1_label_raw"],
            keys_ops=rule["keys_ops"],
        )


# ----------------- GUI ----------------- #

def show_mapping_gui(wb):
    sh_tgt = wb.sheets[TGT_SHEET]

    headers = sh_tgt.range("A1").expand("right").value
    if not isinstance(headers, list):
        headers = [headers]
    sheet1_labels = headers[1:]

    rows_info, code_to_rows, _ = read_feuil1_meta(wb)
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

    items = []   # unified list: (key, op, type)

    lbl_selected = ttk.Label(frame, text="Selected items: (none)")
    lbl_selected.grid(row=8, columnspan=3, sticky="w")

    # ----- codes -----
    ttk.Label(frame, text="3a. Add code:").grid(row=3, sticky="w", pady=(10,0))
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
                f"{op}{key}" for key, op, t in items
            )
            lbl_code_feedback["text"] = "✔ Code added"
            lbl_code_feedback.after(1200, lambda: lbl_code_feedback.config(text=""))

    ttk.Button(frame, text="Add code", command=add_code).grid(row=4, column=1, padx=5)

    # ----- labels -----
    ttk.Label(frame, text="3b. Add label:").grid(row=6, sticky="w", pady=(10,0))
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
                f"{op}{key}" for key, op, t in items
            )
            lbl_lbl_feedback["text"] = "✔ Label added"
            lbl_lbl_feedback.after(1200, lambda: lbl_lbl_feedback.config(text=""))

    ttk.Button(frame, text="Add label", command=add_label).grid(row=7, column=1, padx=5)

    # ----- save checkbox -----
    ttk.Label(frame, text="4. Save this mapping?").grid(row=10, sticky="w", pady=(10,0))
    var_save = tk.BooleanVar(value=True)
    ttk.Checkbutton(frame, text="Save for next time", variable=var_save).grid(row=11, sticky="w")

    # ----- reset -----
    def reset_defaults():
        if not messagebox.askyesno("Reset", "Reset to default?"):
            return

        sh_map = _ensure_custom_sheet(wb)
        sh_map.clear_contents()
        sh_map.range("A1").value = ["Sheet1_Label", "Mapping_Type", "Feuil1_Keys"]

        sh_tgt = wb.sheets[TGT_SHEET]
        last_col = sh_tgt.range("A1").expand("right").columns.count
        last_row = sh_tgt.range("A1").expand("down").rows.count
        sh_tgt.range((2,2),(last_row,last_col)).clear_contents()

        fill_simplified_table()
        wb.save()
        messagebox.showinfo("Reset", "Default restored.")
        root.destroy()

    # ----- apply -----
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

    ttk.Button(frame, text="Reset to default", command=reset_defaults).grid(row=12, column=0, pady=10)
    ttk.Button(frame, text="Apply", command=apply_mapping).grid(row=12, column=1, pady=10)
    ttk.Button(frame, text="Cancel", command=root.destroy).grid(row=12, column=2, pady=10)

    root.mainloop()


# ----------------- Main ----------------- #

def main():
    wb = _get_workbook()
    fill_simplified_table()
    apply_saved_mappings(wb)
    show_mapping_gui(wb)


if __name__ == "__main__":
    wb = _get_workbook()
    fill_simplified_table()
    apply_saved_mappings(wb)
    show_mapping_gui(wb)
