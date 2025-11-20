"""
excel_gui_launcher.py
---------------------

Entry point for the Excel button.

When the button is pressed:
1. Runs the existing automation fill_simplified_table() from excel_automation_script.py
2. Applies any previously saved custom mappings (sheet 'CustomMap')
3. Opens a GUI so the user can define a new mapping:
   - Select a label in Sheet1
   - Map it to Feuil1 rows (by code or by label)
   - Attach an operator to each element: +, -, *, /
   - Decide if the mapping should be saved for future runs or used only once
   - Optional: reset to default filling (delete all custom mappings)
"""

import xlwings as xw
import tkinter as tk
from tkinter import ttk, messagebox

# Import your existing logic and helpers
from excel_automation_script import fill_simplified_table, _norm_label, _norm_emp_id

SRC_SHEET = "Feuil1"       # source sheet with full payroll
TGT_SHEET = "Sheet1"       # target sheet
CUSTOM_SHEET = "CustomMap" # sheet where mappings are stored

BLOCK_WIDTH = 3            # each employee block = 3 columns on Feuil1
MAX_COLS = 120             # max columns to scan for employees


# ----------------- Workbook helpers ----------------- #

def _get_workbook():
    """Attach to the workbook (button or manual run)."""
    try:
        wb = xw.Book.caller()  # works if called from Excel via xlwings addin
    except Exception:
        # Fallback when running via Shell (your case)
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        wb = app.books[0]
    return wb


def _ensure_custom_sheet(wb):
    """Get or create the CustomMap sheet."""
    try:
        sh = wb.sheets[CUSTOM_SHEET]
    except Exception:
        sh = wb.sheets.add(CUSTOM_SHEET)
        sh.range("A1").value = ["Sheet1_Label", "Mapping_Type", "Feuil1_Keys"]
    return sh


# ----------------- Serialization helpers ----------------- #

def serialize_keys_ops(keys_ops):
    """
    keys_ops: list of (key, op), e.g. [("F07", "+"), ("F08", "-")]

    Stored format in Excel cell:
        "+::F07;;-::F08"
    """
    return ";;".join(f"{op}::{key}" for key, op in keys_ops)


def deserialize_keys_ops(cell_value):
    """
    Reverse of serialize_keys_ops → returns list[(key, op)].
    """
    result = []
    if not cell_value:
        return result

    if isinstance(cell_value, list):
        parts = []
        for item in cell_value:
            if item:
                parts.extend(str(item).split(";;"))
    else:
        parts = str(cell_value).split(";;")

    for item in parts:
        item = item.strip()
        if not item:
            continue
        if "::" in item:
            op, key = item.split("::", 1)
        else:
            # Fallback: guess first character as operator
            if item[0] in "+-*/":
                op = item[0]
                key = item[1:].strip()
            else:
                op = "+"
                key = item
        op = op.strip() or "+"
        if op not in ["+", "-", "*", "/"]:
            op = "+"
        key = key.strip()
        if key:
            result.append((key, op))
    return result


# ----------------- CustomMap I/O ----------------- #

def load_saved_mappings(wb):
    """
    Read mappings from the CustomMap sheet.

    Returns a list of rules:
    [
        {
            "sheet1_label_raw": str,
            "sheet1_label_norm": str,
            "mapping_type": "code" or "label",
            "keys_ops": [(key, op), ...],   # e.g. [("F07", "+"), ("F08", "-")]
        },
        ...
    ]
    """
    try:
        sh = wb.sheets[CUSTOM_SHEET]
    except Exception:
        return []

    used = sh.range("A1").current_region
    if used.rows.count <= 1:
        return []

    rows = used.value
    rules = []
    for row in rows[1:]:
        if not row or all(v in [None, ""] for v in row):
            continue

        sheet1_label_raw = row[0]
        mapping_type = (row[1] or "").strip().lower() if len(row) > 1 and row[1] else ""
        keys_cell = row[2] if len(row) > 2 else ""

        if not sheet1_label_raw or mapping_type not in ("code", "label"):
            continue

        keys_ops = deserialize_keys_ops(keys_cell)
        if not keys_ops:
            continue

        rules.append({
            "sheet1_label_raw": str(sheet1_label_raw),
            "sheet1_label_norm": _norm_label(sheet1_label_raw),
            "mapping_type": mapping_type,
            "keys_ops": keys_ops,
        })

    return rules


def save_or_update_mapping_row(wb, sheet1_label_raw, mapping_type, keys_ops):
    """
    Write or update a mapping row in CustomMap.

    - Overwrites existing mapping for the same normalized Sheet1 label (if present).
    - keys_ops is list[(key, op)].
    """
    sh = _ensure_custom_sheet(wb)
    used = sh.range("A1").current_region
    rows = used.rows.count
    target_norm = _norm_label(sheet1_label_raw)

    existing_row_idx = None
    if rows > 1:
        data = used.value
        for i, row in enumerate(data[1:], start=2):
            if not row or row[0] in [None, ""]:
                continue
            if _norm_label(row[0]) == target_norm:
                existing_row_idx = i
                break

    if existing_row_idx is None:
        existing_row_idx = rows + 1

    keys_string = serialize_keys_ops(keys_ops)

    sh.range(existing_row_idx, 1).value = sheet1_label_raw
    sh.range(existing_row_idx, 2).value = mapping_type
    sh.range(existing_row_idx, 3).value = keys_string


# ----------------- Feuil1 & Sheet1 structure ----------------- #

def read_feuil1_meta(wb):
    """
    Robust scanning of Feuil1:
    - Reads ALL used rows in columns A and B
    - Extracts codes (col A) and labels (col B)
    - Skips rows where both are empty
    Returns:
      rows_info: list of dicts:
        {"row_index": int, "code": str, "label_raw": str, "label_norm": str}
      code_to_rows: dict[code] -> [row_indexes...]
      labelnorm_to_rows: dict[label_norm] -> [row_indexes...]
    """
    sh = wb.sheets[SRC_SHEET]

    last_row_a = sh.range("A" + str(sh.cells.rows.count)).end("up").row
    last_row_b = sh.range("B" + str(sh.cells.rows.count)).end("up").row
    last_row = max(last_row_a, last_row_b)

    rows_info = []
    code_to_rows = {}
    labelnorm_to_rows = {}

    for r in range(1, last_row + 1):
        code = sh.range((r, 1)).value
        label_raw = sh.range((r, 2)).value

        if (code is None or str(code).strip() == "") and \
           (label_raw is None or str(label_raw).strip() == ""):
            continue

        code_str = str(code).strip() if code not in [None, ""] else ""
        label_raw_str = str(label_raw).strip() if label_raw not in [None, ""] else ""
        label_norm = _norm_label(label_raw_str) if label_raw_str else ""

        rows_info.append({
            "row_index": r,
            "code": code_str,
            "label_raw": label_raw_str,
            "label_norm": label_norm,
        })

        if code_str:
            code_to_rows.setdefault(code_str, []).append(r)
        if label_norm:
            labelnorm_to_rows.setdefault(label_norm, []).append(r)

    return rows_info, code_to_rows, labelnorm_to_rows


def read_employee_mappings(wb):
    """
    Map employees between Feuil1 and Sheet1 via normalized IDs.

    Returns:
      emp_norm_to_feuil1_col: dict emp_norm -> col_start in Feuil1
      emp_norm_to_sheet1_row: dict emp_norm -> row index in Sheet1
      header_map: dict normalized header (Sheet1) -> column index (1-based)
      id_width: width used for zero-padding IDs
    """
    sh_src = wb.sheets[SRC_SHEET]
    sh_tgt = wb.sheets[TGT_SHEET]

    # Sheet1 headers
    tgt_headers = sh_tgt.range("A1").expand("right").value
    if not isinstance(tgt_headers, list):
        tgt_headers = [tgt_headers]
    header_map = {_norm_label(h): idx + 1 for idx, h in enumerate(tgt_headers)}

    # Sheet1 employee IDs
    tgt_ids = sh_tgt.range("A2").expand("down").value
    if not isinstance(tgt_ids, list):
        tgt_ids = [tgt_ids]
    tgt_ids = [str(v).strip() if v else "" for v in tgt_ids]

    # Determine id_width like in original script
    digits_lengths = [len("".join(filter(str.isdigit, v))) for v in tgt_ids if v]
    id_width = max(digits_lengths) if digits_lengths else 5
    tgt_ids_norm = [_norm_emp_id(v, id_width) for v in tgt_ids]

    emp_norm_to_sheet1_row = {
        emp_norm: i for i, emp_norm in enumerate(tgt_ids_norm, start=2) if emp_norm
    }

    # Feuil1 employee IDs (row 3 across columns)
    row_vals = sh_src.range((3, 1), (3, MAX_COLS)).value
    if not isinstance(row_vals, list):
        row_vals = [row_vals]

    emp_norm_to_feuil1_col = {}
    for col_idx, emp in enumerate(row_vals, start=1):
        if not emp or str(emp).strip() == "":
            continue
        emp_norm = _norm_emp_id(emp, id_width)
        if emp_norm:
            emp_norm_to_feuil1_col[emp_norm] = col_idx

    return emp_norm_to_feuil1_col, emp_norm_to_sheet1_row, header_map, id_width


# ----------------- Core calculation helpers ----------------- #

def _sum_rows_for_employee(sh_src, row_indexes, col_start):
    """
    Sum values for a given employee (block of 3 columns) and list of Feuil1 row indexes.
    """
    total = 0
    for r in row_indexes:
        vals = sh_src.range((r, col_start), (r, col_start + BLOCK_WIDTH - 1)).value
        if isinstance(vals, list):
            for v in vals:
                if isinstance(v, (int, float)):
                    total += v
        else:
            if isinstance(vals, (int, float)):
                total += vals
    return total


def apply_single_mapping(wb, sheet1_label_raw, mapping_type, keys_ops):
    """
    Apply one mapping rule on Sheet1.

    mapping_type: "code" or "label"
    keys_ops: list of (key, op) where op in {"+", "-", "*", "/"}
    """
    sh_src = wb.sheets[SRC_SHEET]
    sh_tgt = wb.sheets[TGT_SHEET]

    # Feuil1 meta
    _, code_to_rows, labelnorm_to_rows = read_feuil1_meta(wb)

    # Employee mappings & headers
    emp_norm_to_feuil1_col, emp_norm_to_sheet1_row, header_map, _ = read_employee_mappings(wb)

    target_norm = _norm_label(sheet1_label_raw)
    if target_norm not in header_map:
        messagebox.showwarning(
            "Mapping",
            f"The label '{sheet1_label_raw}' was not found in Sheet1 headers."
        )
        return

    target_col = header_map[target_norm]

    # For each employee present both in Feuil1 and Sheet1, compute expression
    for emp_norm, col_start in emp_norm_to_feuil1_col.items():
        row_sheet1 = emp_norm_to_sheet1_row.get(emp_norm)
        if not row_sheet1:
            continue

        total = None

        for key, op in keys_ops:
            # Find relevant rows for this key
            if mapping_type == "code":
                code_str = str(key).strip()
                row_indexes = code_to_rows.get(code_str, [])
            else:  # mapping_type == "label"
                label_norm = _norm_label(key)
                row_indexes = labelnorm_to_rows.get(label_norm, [])

            if not row_indexes:
                continue

            part = _sum_rows_for_employee(sh_src, row_indexes, col_start)

            # Initialize accumulator
            if total is None:
                total = 0 if op in ["+", "-"] else 1

            # Apply operator
            if op == "+":
                total += part
            elif op == "-":
                total -= part
            elif op == "*":
                total *= part
            elif op == "/":
                if part != 0:
                    total /= part

        if total is None:
            continue

        cell = sh_tgt.range((row_sheet1, target_col))
        cell.value = total
        cell.color = (204, 255, 204)  # light green
        cell.api.Font.Color = 0
        cell.api.Font.Bold = True


def apply_saved_mappings(wb):
    """Apply all mappings stored in CustomMap sheet."""
    rules = load_saved_mappings(wb)
    if not rules:
        return

    for rule in rules:
        apply_single_mapping(
            wb,
            sheet1_label_raw=rule["sheet1_label_raw"],
            mapping_type=rule["mapping_type"],
            keys_ops=rule["keys_ops"],
        )


# ----------------- GUI ----------------- #

def show_mapping_gui(wb):
    """Launch the Tkinter GUI for defining a new mapping."""
    sh_tgt = wb.sheets[TGT_SHEET]

    # Sheet1 labels (headers), skip column A
    tgt_headers = sh_tgt.range("A1").expand("right").value
    if not isinstance(tgt_headers, list):
        tgt_headers = [tgt_headers]
    sheet1_labels = [h for h in tgt_headers[1:] if h not in [None, ""]]

    # Feuil1 meta
    rows_info, code_to_rows, _ = read_feuil1_meta(wb)
    codes_sorted = sorted([c for c in code_to_rows.keys() if c])

    # Build label display list (label + optional code)
    labels_display = []
    label_display_to_raw = {}
    for info in rows_info:
        label_raw = info["label_raw"]
        code = info["code"]
        if not label_raw:
            continue
        if code:
            display = f"{label_raw} ({code})"
        else:
            display = label_raw
        labels_display.append(display)
        label_display_to_raw[display] = label_raw

    labels_display = sorted(set(labels_display))

    # ----- Tkinter UI ----- #

    root = tk.Tk()
    root.title("Custom mapping between Sheet1 and Feuil1")

    frame = ttk.Frame(root, padding=10)
    frame.grid(row=0, column=0, sticky="nsew")

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # 1. Select label from Sheet1
    ttk.Label(frame, text="1. Select a label from Sheet1:").grid(row=0, column=0, columnspan=4, sticky="w")

    sheet1_label_var = tk.StringVar()
    cb_sheet1 = ttk.Combobox(frame, textvariable=sheet1_label_var, state="readonly")
    cb_sheet1["values"] = sheet1_labels
    cb_sheet1.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(0, 10))

    # 2. Mapping type
    ttk.Label(frame, text="2. Choose how to map from Feuil1:").grid(row=2, column=0, columnspan=4, sticky="w", pady=(5, 0))

    mapping_type_var = tk.StringVar(value="code")
    rb_code = ttk.Radiobutton(frame, text="By code (column A of Feuil1)", variable=mapping_type_var, value="code")
    rb_label = ttk.Radiobutton(frame, text="By label name (column B of Feuil1)", variable=mapping_type_var, value="label")
    rb_code.grid(row=3, column=0, columnspan=4, sticky="w")
    rb_label.grid(row=4, column=0, columnspan=4, sticky="w")

    # 3. Operator selection
    ttk.Label(frame, text="Operator for next item:").grid(row=5, column=0, sticky="w", pady=(10, 0))
    operator_var = tk.StringVar(value="+")
    cb_operator = ttk.Combobox(frame, textvariable=operator_var, state="readonly", width=5)
    cb_operator["values"] = ["+", "-", "*", "/"]
    cb_operator.grid(row=5, column=1, sticky="w", pady=(10, 0))

    # 3a. Codes
    ttk.Label(frame, text="3a. Select Feuil1 codes and click 'Add':").grid(row=6, column=0, columnspan=4, sticky="w", pady=(10, 0))
    code_var = tk.StringVar()
    cb_code = ttk.Combobox(frame, textvariable=code_var, state="readonly")
    cb_code["values"] = codes_sorted
    cb_code.grid(row=7, column=0, columnspan=2, sticky="ew")

    selected_codes = []  # list[(code, op)]

    lbl_selected_codes = ttk.Label(frame, text="Selected codes: (none)")
    lbl_selected_codes.grid(row=8, column=0, columnspan=4, sticky="w")

    lbl_code_feedback = ttk.Label(frame, text="", foreground="green")
    lbl_code_feedback.grid(row=9, column=0, columnspan=4, sticky="w")

    def add_code():
        val = code_var.get()
        if not val:
            return
        op = operator_var.get() or "+"
        pair = (val, op)
        if pair not in selected_codes:
            selected_codes.append(pair)
            disp = ", ".join(f"{op}{code}" for code, op in selected_codes)
            lbl_selected_codes["text"] = "Selected codes: " + disp
            lbl_code_feedback["text"] = "✔ Code added"
            lbl_code_feedback.after(1200, lambda: lbl_code_feedback.config(text=""))

    btn_add_code = ttk.Button(frame, text="Add code", command=add_code)
    btn_add_code.grid(row=7, column=2, padx=5, sticky="w")

    # 3b. Labels
    ttk.Label(frame, text="3b. OR select Feuil1 labels and click 'Add':").grid(row=10, column=0, columnspan=4, sticky="w", pady=(10, 0))
    label_var = tk.StringVar()
    cb_label = ttk.Combobox(frame, textvariable=label_var, state="readonly")
    cb_label["values"] = labels_display
    cb_label.grid(row=11, column=0, columnspan=2, sticky="ew")

    selected_labels = []  # list[(label_raw, op)]

    lbl_selected_labels = ttk.Label(frame, text="Selected labels: (none)")
    lbl_selected_labels.grid(row=12, column=0, columnspan=4, sticky="w")

    lbl_label_feedback = ttk.Label(frame, text="", foreground="green")
    lbl_label_feedback.grid(row=13, column=0, columnspan=4, sticky="w")

    def add_label():
        display = label_var.get()
        if not display:
            return
        raw = label_display_to_raw.get(display, display)
        op = operator_var.get() or "+"
        pair = (raw, op)
        if pair not in selected_labels:
            selected_labels.append(pair)
            disp = ", ".join(f"{op}{lbl}" for lbl, op in selected_labels)
            lbl_selected_labels["text"] = "Selected labels: " + disp
            lbl_label_feedback["text"] = "✔ Label added"
            lbl_label_feedback.after(1200, lambda: lbl_label_feedback.config(text=""))

    btn_add_label = ttk.Button(frame, text="Add label", command=add_label)
    btn_add_label.grid(row=11, column=2, padx=5, sticky="w")

    # 4. Save mapping?
    ttk.Label(frame, text="4. Save this mapping?").grid(row=14, column=0, columnspan=4, sticky="w", pady=(10, 0))
    save_for_future_var = tk.BooleanVar(value=True)
    chk_save = ttk.Checkbutton(frame, text="Save and reuse this mapping next time", variable=save_for_future_var)
    chk_save.grid(row=15, column=0, columnspan=4, sticky="w")

    # Reset to default
    def reset_defaults():
        if not messagebox.askyesno(
            "Reset to default",
            "This will delete ALL custom mappings and re-run the default filling.\n"
            "Are you sure?"
        ):
            return
        sh = _ensure_custom_sheet(wb)
        sh.clear_contents()
        sh.range("A1").value = ["Sheet1_Label", "Mapping_Type", "Feuil1_Keys"]
        wb.save()

        # Re-run your original fill without any custom mapping
        fill_simplified_table()
        messagebox.showinfo("Reset", "Custom mappings cleared and default filling reapplied.")
        root.destroy()

    # Apply / Cancel
    def on_apply():
        sheet1_label = sheet1_label_var.get()
        if not sheet1_label:
            messagebox.showerror("Error", "Please select a label from Sheet1.")
            return

        mtype = mapping_type_var.get()
        if mtype == "code":
            keys_ops = selected_codes.copy()
            if not keys_ops:
                messagebox.showerror("Error", "Please add at least one Feuil1 code.")
                return
        else:
            keys_ops = selected_labels.copy()
            if not keys_ops:
                messagebox.showerror("Error", "Please add at least one Feuil1 label.")
                return

        # Apply now
        apply_single_mapping(wb, sheet1_label_raw=sheet1_label, mapping_type=mtype, keys_ops=keys_ops)

        # Save if requested
        if save_for_future_var.get():
            save_or_update_mapping_row(wb, sheet1_label_raw=sheet1_label, mapping_type=mtype, keys_ops=keys_ops)
            wb.save()
            messagebox.showinfo("Mapping", "Mapping applied and saved for future runs.")
        else:
            messagebox.showinfo("Mapping", "Mapping applied only for this run (not saved).")

        root.destroy()

    def on_cancel():
        root.destroy()

    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=16, column=0, columnspan=4, pady=(15, 0), sticky="e")

    btn_apply = ttk.Button(btn_frame, text="Apply mapping", command=on_apply)
    btn_cancel = ttk.Button(btn_frame, text="Cancel", command=on_cancel)
    btn_reset = ttk.Button(btn_frame, text="Reset to default", command=reset_defaults)

    btn_reset.grid(row=0, column=0, padx=5)
    btn_apply.grid(row=0, column=1, padx=5)
    btn_cancel.grid(row=0, column=2, padx=5)

    root.mainloop()


# ----------------- Main entry for Excel button ----------------- #

def main():
    """
    Entry function to be called from the Excel button.

    1. Run existing automation (fill_simplified_table)
    2. Apply previously saved mappings from CustomMap
    3. Open GUI so the user can define a new mapping
    """
    wb = _get_workbook()

    # Step 1: run your original logic
    fill_simplified_table()

    # Step 2: apply any saved custom mappings
    apply_saved_mappings(wb)

    # Step 3: launch GUI for new mapping
    show_mapping_gui(wb)

    try:
        wb.app.api.StatusBar = "✅ Data updated + custom mappings applied."
    except Exception:
        pass


# For manual testing (outside Excel)
if __name__ == "__main__":
    wb = _get_workbook()
    fill_simplified_table()
    apply_saved_mappings(wb)
    show_mapping_gui(wb)
