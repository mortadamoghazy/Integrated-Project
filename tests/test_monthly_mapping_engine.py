"""
tests/test_monthly_mapping_engine.py

Multi-month, user-defined mapping engine (NO hard-coded payroll logic).

What it does:
- Detect month source sheets by regex (supports YYYY-MM OR MM).
- For each month source sheet:
  - Create a new destination sheet named "<month>_mapped" (configurable suffix).
  - Write the destination header template (based on your screenshot, WITHOUT "role").
  - Populate column A (employee_id) from the source sheet's employee ID row.
  - Apply saved mapping rules from CustomMap to fill other columns.

Mapping rules:
- Stored in source workbook sheet "CustomMap"
- Columns:
  A: Destination_Label (must match destination header label)
  B: Mapping_Type (kept for compatibility; not used)
  C: KeysOps string "op::key::type;;op::key::type..."
     op in {+,-,*,/}, type in {"code","label"}
     key is either a row code (col A in source) or a row label (col B in source)
"""

from __future__ import annotations

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import re
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional

import xlwings as xw

from src.core.normalization import _norm_label, _norm_emp_id


# ---------------------------------------------------------------------
# FIX_ME: SOURCE AND OUTPUT EXCEL PATHS
# ---------------------------------------------------------------------
# Update these paths to match your local file system structure
# SOURCE_EXCEL_PATH: The raw payroll Excel file to be processed
# OUTPUT_EXCEL_PATH: Where the mapped/sorted data will be saved
#
# IMPORTANT: Use raw strings (r"") for Windows paths to avoid escape issues
# Example: r"C:\Users\YourName\Documents\Project\data\raw\payroll.xlsx"
# ---------------------------------------------------------------------
SOURCE_EXCEL_PATH = r"C:\Users\Morta\OneDrive\Desktop\Gam3a\MARS\Integrated Project\data\raw\essai donnees paye.xlsx"
OUTPUT_EXCEL_PATH = r"C:\Users\Morta\OneDrive\Desktop\Gam3a\MARS\Integrated Project\data\raw\data_sorted.xlsx"


# ---------------------------------------------------------------------
# DESTINATION TEMPLATE (fixed structure, no formulas)
# Matches your screenshot order, but with "role" removed.
# ---------------------------------------------------------------------
DEST_HEADERS = [
    "employee_id",
    "salaire_brut",
    "cot_salariale",
    "cot_patronale",
    "net_imposable",
    "PAS",
    "net_paye",
    "avantages",
    "total_cost",
]


# ---------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------
CUSTOM_SHEET = "CustomMap"

# IMPORTANT FIX:
# Supports either:
#   - YYYY-MM  (e.g., 2023-01)
#   - MM       (e.g., 06, 11)
DEFAULT_MONTH_SHEET_REGEX = r"^(?:\d{4}-(?:0[1-9]|1[0-2])|(?:0[1-9]|1[0-2]))$"

DEFAULT_OUTPUT_SUFFIX = "_mapped"

# Source layout assumptions (structural, not payroll logic)
DEFAULT_ID_ROW = 3       # employee IDs row in each month sheet
DEFAULT_LABEL_COL = 2    # column B contains labels
DEFAULT_CODE_COL = 1     # column A contains codes


# ---------------------------------------------------------------------
# Mapping serialization helpers (compatible with your current format)
# ---------------------------------------------------------------------
def deserialize_keys_ops(cell_value: str) -> List[Tuple[str, str, str]]:
    """
    Parse: "op::key::type;;op::key::type"
    Return: [(key, op, type), ...]
    """
    result: List[Tuple[str, str, str]] = []
    if not cell_value:
        return result
    for entry in str(cell_value).split(";;"):
        if "::" not in entry:
            continue
        try:
            op, key, item_type = entry.split("::")
        except Exception:
            continue
        op = op.strip()
        key = key.strip()
        item_type = item_type.strip().lower()
        if op in ["+", "-", "*", "/"] and key and item_type in ["code", "label"]:
            result.append((key, op, item_type))
    return result


@dataclass
class MappingRule:
    dest_label_raw: str
    dest_label_norm: str
    keys_ops: List[Tuple[str, str, str]]  # (key, op, type)


def ensure_custom_sheet(wb: xw.Book) -> xw.Sheet:
    try:
        sh = wb.sheets[CUSTOM_SHEET]
    except Exception:
        sh = wb.sheets.add(CUSTOM_SHEET)
        sh.range("A1").value = ["Destination_Label", "Mapping_Type", "KeysOps"]
    return sh


def load_saved_mappings(wb: xw.Book) -> List[MappingRule]:
    """
    Read mappings from CustomMap:
    - A: Destination_Label
    - C: KeysOps serialized string
    """
    try:
        sh = wb.sheets[CUSTOM_SHEET]
    except Exception:
        return []

    used = sh.range("A1").current_region
    if used.rows.count <= 1:
        return []

    rules: List[MappingRule] = []
    for row in used.value[1:]:
        if not row or not row[0]:
            continue
        dest_raw = str(row[0]).strip()
        keys_ops = deserialize_keys_ops(row[2] if len(row) > 2 else "")
        if not keys_ops:
            continue
        rules.append(
            MappingRule(
                dest_label_raw=dest_raw,
                dest_label_norm=_norm_label(dest_raw),
                keys_ops=keys_ops,
            )
        )
    return rules


# ---------------------------------------------------------------------
# Month sheet detection (FIXED)
# ---------------------------------------------------------------------
def is_month_sheet(name: str, month_regex: str) -> bool:
    return re.match(month_regex, name.strip()) is not None


def list_month_source_sheets(
    wb: xw.Book,
    month_regex: str = DEFAULT_MONTH_SHEET_REGEX,
    output_suffix: str = DEFAULT_OUTPUT_SUFFIX,
    exclude_sheets: Optional[List[str]] = None,
) -> List[str]:
    """
    Returns month sheets only, excluding:
    - CustomMap
    - Any sheet already generated (ending with output_suffix)
    - Any additional exclude_sheets provided
    """
    exclude = {CUSTOM_SHEET}
    if exclude_sheets:
        exclude.update(exclude_sheets)

    out: List[str] = []
    for sh in wb.sheets:
        name = sh.name

        if name in exclude:
            continue
        if name.endswith(output_suffix):
            continue
        if is_month_sheet(name, month_regex):
            out.append(name)

    out.sort()
    return out


# ---------------------------------------------------------------------
# Destination sheet creation
# ---------------------------------------------------------------------
def create_destination_sheet(
    wb: xw.Book,
    base_name: str,
    suffix: str = DEFAULT_OUTPUT_SUFFIX,
) -> xw.Sheet:
    """
    Create a new destination sheet with the same name as source (no suffix).
    """
    name = base_name
    existing = {s.name for s in wb.sheets}
    k = 1
    while name in existing:
        name = f"{base_name}_{k}"
        k += 1

    sh = wb.sheets.add(name, after=wb.sheets[-1])

    # Write headers
    sh.range((1, 1), (1, len(DEST_HEADERS))).value = DEST_HEADERS

    # Simple header formatting
    header_rng = sh.range((1, 1), (1, len(DEST_HEADERS)))
    header_rng.api.Font.Bold = True
    header_rng.color = (220, 230, 241)

    # Freeze header (make sure it targets this sheet window)
    sh.activate()
    sh.api.Application.ActiveWindow.SplitRow = 1
    sh.api.Application.ActiveWindow.FreezePanes = True

    # Column widths
    sh.range((1, 1)).column_width = 14
    for idx in range(2, len(DEST_HEADERS) + 1):
        sh.range((1, idx)).column_width = 18

    return sh


def get_dest_header_map() -> Dict[str, int]:
    """
    Normalized header -> column index (1-based) for DEST_HEADERS.
    """
    return {_norm_label(h): i + 1 for i, h in enumerate(DEST_HEADERS)}


# ---------------------------------------------------------------------
# Source parsing (generic)
# ---------------------------------------------------------------------
def read_employee_ids_from_source(sh_src: xw.Sheet, id_row: int) -> Tuple[List[str], int]:
    """
    Read employee IDs across the id_row. Returns (raw_ids_in_order, id_width).
    """
    last_col = sh_src.range(id_row, sh_src.cells.columns.count).end("left").column
    row_vals = sh_src.range((id_row, 1), (id_row, last_col)).value
    if not isinstance(row_vals, list):
        row_vals = [row_vals]

    raw_ids = [
        str(v).strip()
        for v in row_vals
        if v not in (None, "") and str(v).strip() != ""
    ]
    digit_lengths = [len("".join(ch for ch in v if ch.isdigit())) for v in raw_ids if v]
    id_width = max(digit_lengths) if digit_lengths else 5
    return raw_ids, id_width


def detect_employee_blocks_on_source(sh_src: xw.Sheet, id_row: int, id_width: int) -> Dict[str, Dict[str, int]]:
    """
    Detect employee blocks by scanning the id_row:
    - each non-empty cell starts a new employee block
    """
    last_col = sh_src.range(id_row, sh_src.cells.columns.count).end("left").column
    row_vals = sh_src.range((id_row, 1), (id_row, last_col)).value
    if not isinstance(row_vals, list):
        row_vals = [row_vals]

    emp_blocks: Dict[str, Dict[str, int]] = {}
    current_emp = None

    for col_idx, raw in enumerate(row_vals, start=1):
        if raw not in (None, "") and str(raw).strip() != "":
            emp_norm = _norm_emp_id(raw, id_width)

            if current_emp is not None:
                emp_blocks[current_emp]["end_col"] = col_idx - 1

            emp_blocks[emp_norm] = {"start_col": col_idx, "end_col": None}
            current_emp = emp_norm

    if current_emp is not None and emp_blocks[current_emp]["end_col"] is None:
        emp_blocks[current_emp]["end_col"] = last_col

    return emp_blocks


def build_row_lookup(
    sh_src: xw.Sheet,
    code_col: int,
    label_col: int
) -> Tuple[Dict[str, List[int]], Dict[str, List[int]]]:
    """
    Build:
    - code_to_rows: code string -> [row indices]
    - labelnorm_to_rows: normalized label -> [row indices]
    """
    last_row = max(
        sh_src.range((sh_src.cells.rows.count, code_col)).end("up").row,
        sh_src.range((sh_src.cells.rows.count, label_col)).end("up").row,
    )

    code_to_rows: Dict[str, List[int]] = {}
    labelnorm_to_rows: Dict[str, List[int]] = {}

    for r in range(1, last_row + 1):
        code = sh_src.range((r, code_col)).value
        label = sh_src.range((r, label_col)).value
        if not code and not label:
            continue

        if code:
            code_str = str(code).strip()
            if code_str:
                code_to_rows.setdefault(code_str, []).append(r)

        if label:
            lab_str = str(label).strip()
            if lab_str:
                labelnorm_to_rows.setdefault(_norm_label(lab_str), []).append(r)

    return code_to_rows, labelnorm_to_rows


def sum_rows_for_block(sh_src: xw.Sheet, rows: List[int], start_col: int, end_col: int) -> float:
    """
    Sum numeric values across given rows and source block columns.
    """
    total = 0.0
    for r in rows:
        vals = sh_src.range((r, start_col), (r, end_col)).value
        if isinstance(vals, list):
            for v in vals:
                if isinstance(v, (int, float)):
                    total += float(v)
        else:
            if isinstance(vals, (int, float)):
                total += float(vals)
    return total


# ---------------------------------------------------------------------
# Apply mappings for one month
# ---------------------------------------------------------------------
def fill_dest_employee_ids(sh_dest: xw.Sheet, raw_ids: List[str], id_width: int) -> Dict[str, int]:
    """
    Write employee_id down column A starting at row 2.
    Returns emp_norm -> destination row.
    """
    if not raw_ids:
        return {}

    sh_dest.range((2, 1), (1 + len(raw_ids), 1)).value = [[rid] for rid in raw_ids]

    emp_to_row: Dict[str, int] = {}
    for idx, rid in enumerate(raw_ids, start=2):
        emp_norm = _norm_emp_id(rid, id_width)
        emp_to_row[emp_norm] = idx
    return emp_to_row


def apply_mappings_one_month(
    sh_src: xw.Sheet,
    sh_dest: xw.Sheet,
    rules: List[MappingRule],
    *,
    id_row: int,
    code_col: int,
    label_col: int,
) -> None:
    dest_header_map = get_dest_header_map()

    raw_ids, id_width = read_employee_ids_from_source(sh_src, id_row=id_row)
    emp_to_dest_row = fill_dest_employee_ids(sh_dest, raw_ids, id_width)
    emp_blocks = detect_employee_blocks_on_source(sh_src, id_row=id_row, id_width=id_width)

    code_to_rows, labelnorm_to_rows = build_row_lookup(sh_src, code_col=code_col, label_col=label_col)

    for rule in rules:
        dest_col = dest_header_map.get(rule.dest_label_norm)
        if not dest_col or dest_col == 1:
            continue  # skip employee_id col or unknown headers

        for emp_norm, block in emp_blocks.items():
            dest_row = emp_to_dest_row.get(emp_norm)
            if not dest_row:
                continue

            start_col, end_col = block["start_col"], block["end_col"]
            total: Optional[float] = None

            for key, op, item_type in rule.keys_ops:
                if item_type == "code":
                    rows = code_to_rows.get(key, [])
                else:
                    rows = labelnorm_to_rows.get(_norm_label(key), [])

                if not rows:
                    continue

                part = sum_rows_for_block(sh_src, rows, start_col, end_col)

                if total is None:
                    total = 0.0 if op in ["+", "-"] else 1.0

                if op == "+":
                    total += part
                elif op == "-":
                    total -= part
                elif op == "*":
                    total *= part
                elif op == "/" and part != 0:
                    total /= part

            if total is not None:
                cell = sh_dest.range((dest_row, dest_col))
                cell.value = total
                cell.color = (204, 255, 204)
                cell.api.Font.Bold = True


# ---------------------------------------------------------------------
# Main runner
# ---------------------------------------------------------------------
def run_all_months(
    *,
    month_regex: str = DEFAULT_MONTH_SHEET_REGEX,
    output_suffix: str = DEFAULT_OUTPUT_SUFFIX,
    id_row: int = DEFAULT_ID_ROW,
    code_col: int = DEFAULT_CODE_COL,
    label_col: int = DEFAULT_LABEL_COL,
    exclude_sheets: Optional[List[str]] = None,
) -> None:
    """
    Entry point:
    - open SOURCE_EXCEL_PATH
    - load mappings once
    - process each month sheet and create output in new workbook
    """
    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)

    # Reuse workbook if already open
    wb_src = None
    source_name = os.path.basename(SOURCE_EXCEL_PATH)
    
    for b in app.books:
        try:
            if b.name.lower() == source_name.lower():
                wb_src = b
                break
        except:
            pass
    
    if wb_src is None:
        wb_src = app.books.open(SOURCE_EXCEL_PATH)

    print("✅ Using source workbook:", wb_src.name)

    # Check for mapping sheet in source
    if CUSTOM_SHEET not in [s.name for s in wb_src.sheets]:
        raise RuntimeError(
            f"No '{CUSTOM_SHEET}' sheet found in source workbook. "
            "Run tests/test_monthly_mapping_gui_launcher.py first to define mappings."
        )

    rules = load_saved_mappings(wb_src)
    if not rules:
        raise RuntimeError(
            "No saved mappings found in mapping sheet. "
            "Run tests/test_monthly_mapping_gui_launcher.py first to define mappings."
        )

    month_sheets = list_month_source_sheets(
        wb_src,
        month_regex=month_regex,
        output_suffix=output_suffix,
        exclude_sheets=exclude_sheets,
    )
    if not month_sheets:
        raise RuntimeError(
            f"No month sheets matched regex: {month_regex}\n"
            "Tip: your month sheets appear to be named like '06', '07', etc. "
            "This script now supports that, so check the sheet names and regex."
        )

    print(f"Found {len(month_sheets)} month sheets:", month_sheets)

    # Create or reuse output workbook if already open
    output_name = os.path.basename(OUTPUT_EXCEL_PATH)
    wb_out = None
    for b in app.books:
        try:
            if b.name.lower() == output_name.lower():
                wb_out = b
                break
        except Exception:
            pass
    if wb_out is None:
        wb_out = app.books.add()

    for src_name in month_sheets:
        # If a sheet with the same name already exists in the output, delete it to avoid duplicates
        existing_out_names = [s.name for s in wb_out.sheets]
        if src_name in existing_out_names:
            wb_out.sheets[src_name].delete()
        sh_src = wb_src.sheets[src_name]
        sh_dest = create_destination_sheet(wb_out, base_name=src_name, suffix="")

        print(f"Mapping source '{src_name}' -> destination '{sh_dest.name}'")
        apply_mappings_one_month(
            sh_src=sh_src,
            sh_dest=sh_dest,
            rules=rules,
            id_row=id_row,
            code_col=code_col,
            label_col=label_col,
        )

    # Save output workbook (avoid SaveAs conflict if it's already the same file)
    try:
        if hasattr(wb_out, "fullname") and isinstance(wb_out.fullname, str) and wb_out.fullname:
            if os.path.abspath(wb_out.fullname).lower() == os.path.abspath(OUTPUT_EXCEL_PATH).lower():
                wb_out.save()
            else:
                wb_out.save(OUTPUT_EXCEL_PATH)
        else:
            wb_out.save(OUTPUT_EXCEL_PATH)
    except Exception:
        # Fallback: try SaveAs with a suffix
        base, ext = os.path.splitext(OUTPUT_EXCEL_PATH)
        wb_out.save(base + "_1" + ext)
    wb_out.app.calculate()
    print(f"✅ Completed: output saved to {OUTPUT_EXCEL_PATH}")


if __name__ == "__main__":
    run_all_months()
