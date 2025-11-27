"""
mapping_engine.py
Uses data_loader for Feuil1 and employee mappings
and focuses only on applying user-defined mappings.
"""

import xlwings as xw

from src.core.config import SRC_SHEET, TGT_SHEET, CUSTOM_SHEET
from src.core.normalization import _norm_label
from src.core.data_loader import (
    get_feuil1_rows_info,
    get_employee_layout_and_sheet1,
)


# ----------------- Workbook / sheet helpers ----------------- #

def _ensure_custom_sheet(wb: xw.Book):
    try:
        sh = wb.sheets[CUSTOM_SHEET]
    except Exception:
        sh = wb.sheets.add(CUSTOM_SHEET)
        sh.range("A1").value = ["Sheet1_Label", "Mapping_Type", "Feuil1_Keys"]
    return sh


# ----------------- Serialization ----------------- #

def serialize_keys_ops(items):
    return ";;".join(f"{op}::{key}::{item_type}" for key, op, item_type in items)


def deserialize_keys_ops(cell_value):
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
        keys_ops = deserialize_keys_ops(row[2])

        if not keys_ops:
            continue

        rules.append({
            "sheet1_label_raw": sheet1_label,
            "sheet1_label_norm": _norm_label(sheet1_label),
            "keys_ops": keys_ops,
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


# ----------------- Summation helper ----------------- #

def _sum_rows_for_employee(sh_src, rows, start_col, end_col):
    """
    Sum all numeric values in given 'rows' and column range [start_col, end_col]
    for a single employee block.
    """
    total = 0
    for r in rows:
        vals = sh_src.range((r, start_col), (r, end_col)).value
        if isinstance(vals, list):
            for v in vals:
                if isinstance(v, (int, float)):
                    total += v
        else:
            if isinstance(vals, (int, float)):
                total += vals
    return total


# ----------------- Apply mapping ----------------- #

def apply_single_mapping(wb, sheet1_label_raw, keys_ops):
    sh_src = wb.sheets[SRC_SHEET]
    sh_tgt = wb.sheets[TGT_SHEET]

    _, code_to_rows, labelnorm_to_rows = get_feuil1_rows_info(wb)
    emp_to_f1, emp_to_s1, header_map = get_employee_layout_and_sheet1(wb)

    target_col = header_map.get(_norm_label(sheet1_label_raw))
    if not target_col:
        return

    for emp_norm, (start_col, end_col) in emp_to_f1.items():
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

            part = _sum_rows_for_employee(sh_src, rows, start_col, end_col)

            if total is None:
                total = 0 if op in ["+", "-"] else 1

            if op == "+":
                total += part
            elif op == "-":
                total -= part
            elif op == "*":
                total *= part
            elif op == "/" and part != 0:
                total /= part

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


def reset_to_default(wb):
    sh_map = _ensure_custom_sheet(wb)
    sh_map.clear_contents()
    sh_map.range("A1").value = ["Sheet1_Label", "Mapping_Type", "Feuil1_Keys"]

    sh_tgt = wb.sheets[TGT_SHEET]
    last_col = sh_tgt.range("A1").expand("right").columns.count
    last_row = sh_tgt.range("A1").expand("down").rows.count
    sh_tgt.range((2, 2), (last_row, last_col)).clear_contents()
