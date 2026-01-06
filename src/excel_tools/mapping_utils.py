"""
mapping_utils.py

Shared utility functions for Excel mapping operations.
Used by both import_dialog.py and test_monthly_mapping_gui_launcher.py
to avoid code duplication.
"""

from __future__ import annotations

import re
from typing import List, Tuple, Dict

import xlwings as xw

from src.core.normalization import _norm_label


CUSTOM_SHEET = "CustomMap"
DEFAULT_MONTH_SHEET_REGEX = r"^(?:\d{4}-(?:0[1-9]|1[0-2])|(?:0[1-9]|1[0-2]))$"


def ensure_custom_sheet(wb: xw.Book) -> xw.Sheet:
    """Ensure the CustomMap sheet exists in the workbook with proper headers."""
    if CUSTOM_SHEET not in [s.name for s in wb.sheets]:
        sh = wb.sheets.add(CUSTOM_SHEET)
        sh.range("A1").value = ["Destination_Label", "Mapping_Type", "KeysOps"]
    else:
        sh = wb.sheets[CUSTOM_SHEET]
    return sh


def serialize_keys_ops(items: List[Tuple[str, str, str]]) -> str:
    """Serialize list of (key, op, type) tuples into a string format."""
    return ";;".join(f"{op}::{key}::{item_type}" for key, op, item_type in items)


def is_month_sheet(name: str, month_regex: str) -> bool:
    """Check if a sheet name matches the month pattern."""
    return re.match(month_regex, name.strip()) is not None


def get_month_sheets(wb: xw.Book, month_regex: str) -> List[str]:
    """Get all sheets in workbook that match the month regex pattern."""
    names = []
    for sh in wb.sheets:
        if sh.name == CUSTOM_SHEET:
            continue
        if is_month_sheet(sh.name, month_regex):
            names.append(sh.name)
    names.sort()
    return names


def get_source_rows_display(sh_src: xw.Sheet, label_col: int = 2) -> Tuple[List[str], Dict[str, str], List[str]]:
    """
    Extract displayable row information from a source sheet.
    
    Returns:
    - labels_display: list like "Salaire de base (A12)" or "Label only"
    - display_to_raw_label: mapping back to raw label
    - codes_sorted: sorted list of unique codes
    """
    try:
        last_row = sh_src.used_range.last_cell.row
    except Exception:
        last_row = max(
            sh_src.range("A" + str(sh_src.cells.rows.count)).end("up").row,
            sh_src.range(sh_src.cells.rows.count, label_col).end("up").row,
        )

    rows_info = []
    code_to_rows = {}

    for r in range(1, last_row + 1):
        code = sh_src.range((r, 1)).value
        label_raw = sh_src.range((r, label_col)).value
        if not code and not label_raw:
            continue

        if code is None:
            code_str = ""
        else:
            code_str = str(code).strip()
        label_str = str(label_raw).strip() if label_raw else ""

        rows_info.append((label_str, code_str))
        if code_str:
            code_to_rows.setdefault(code_str, []).append(r)

    codes_sorted = sorted(code_to_rows.keys())

    labels_display = []
    display_to_raw = {}
    for label_str, code_str in rows_info:
        if not label_str:
            continue
        disp = f"{label_str} ({code_str})" if code_str else label_str
        labels_display.append(disp)
        display_to_raw[disp] = label_str

    labels_display = sorted(set(labels_display))
    return labels_display, display_to_raw, codes_sorted


def save_mapping_row(wb: xw.Book, dest_label_raw: str, keys_ops: List[Tuple[str, str, str]]) -> None:
    """Save or update a mapping rule in the CustomMap sheet."""
    sh = ensure_custom_sheet(wb)
    used = sh.range("A1").current_region
    rows = used.rows.count

    target_norm = _norm_label(dest_label_raw)
    existing_row = None

    if rows > 1:
        data = used.value
        for idx, row in enumerate(data[1:], start=2):
            if row and row[0] and _norm_label(row[0]) == target_norm:
                existing_row = idx
                break

    if existing_row is None:
        existing_row = rows + 1

    sh.range(existing_row, 1).value = dest_label_raw
    sh.range(existing_row, 2).value = "mixed"
    sh.range(existing_row, 3).value = serialize_keys_ops(keys_ops)
