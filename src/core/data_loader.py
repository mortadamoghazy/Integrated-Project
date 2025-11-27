"""
data_loader.py
Centralized helpers for reading data from the Excel workbook.

Key properties:
- Employee IDs are detected by scanning a given row (default: row 3).
- Field labels are detected by scanning a given column (default: column B) from a start row (default: row 5).
- Employee "blocks" (set of columns belonging to one employee) are deduced
  by scanning horizontally from one non-empty ID to the next.
- No hard-coded block width: width is computed dynamically per workbook.
"""

import xlwings as xw

from src.core.config import SRC_SHEET, TGT_SHEET
from src.core.normalization import _norm_label, _norm_emp_id

# Defaults for current template (later we can make these user-configurable via GUI)
DEFAULT_ID_ROW = 3
DEFAULT_FIELD_COL = 2          # column B
DEFAULT_FIELD_ROW_START = 5    # first field label row


# ---------------------------------------------------------------------------
# 1. Sheet1 metadata
# ---------------------------------------------------------------------------

def get_sheet1_meta(wb, tgt_sheet_name: str = TGT_SHEET):
    """
    Read Sheet1 header row and employee IDs.

    Returns dict:
      - sheet
      - headers          (raw labels, list)
      - headers_norm     (normalized labels)
      - header_map       (normalized label -> column number, 1-based)
      - ids              (raw employee IDs as strings)
      - ids_norm         (normalized / zero-padded IDs)
      - id_width         (number of digits used for normalization)
    """
    sh_tgt = wb.sheets[tgt_sheet_name]

    headers = sh_tgt.range("A1").expand("right").value
    if not isinstance(headers, list):
        headers = [headers]
    headers_norm = [_norm_label(h) for h in headers]

    ids = sh_tgt.range("A2").expand("down").value
    if not isinstance(ids, list):
        ids = [ids]
    ids = [str(v).strip() if v else "" for v in ids]

    digit_lengths = [len("".join(filter(str.isdigit, v))) for v in ids if v]
    id_width = max(digit_lengths) if digit_lengths else 5
    ids_norm = [_norm_emp_id(v, id_width) for v in ids]

    header_map = {lab: idx + 1 for idx, lab in enumerate(headers_norm)}

    return {
        "sheet": sh_tgt,
        "headers": headers,
        "headers_norm": headers_norm,
        "header_map": header_map,
        "ids": ids,
        "ids_norm": ids_norm,
        "id_width": id_width,
    }


# ---------------------------------------------------------------------------
# 2. Feuil1: employee blocks + row metadata
# ---------------------------------------------------------------------------

def detect_employee_blocks(
    wb,
    id_row: int = DEFAULT_ID_ROW,
):
    """
    Scan the ID row on Feuil1 (default row 3) to detect:
      - which columns belong to which employee
      - where each employee block starts and ends

    Logic:
      - Scan row 'id_row' from column 1 to last used column.
      - Each non-empty cell is treated as a new employee ID start.
      - The block for an employee extends from its start column up to
        (just before) the next employee start column.
      - The last employee's block extends until the last used column.

    Returns:
      emp_blocks: dict[emp_norm] = {
          "raw_id": original cell contents,
          "start_col": first column of block (1-based),
          "end_col": last column of block (1-based),
      }
      id_width: number of digits used for normalization
    """
    sh_src = wb.sheets[SRC_SHEET]

    # Determine last used column in the ID row
    last_col = sh_src.range(id_row, sh_src.cells.columns.count).end("left").column
    row_vals = sh_src.range((id_row, 1), (id_row, last_col)).value
    if not isinstance(row_vals, list):
        row_vals = [row_vals]

    # First pass: collect all raw IDs (for id_width determination)
    raw_ids = [v for v in row_vals if v not in (None, "")]
    raw_ids_str = [str(v) for v in raw_ids]

    digit_lengths = [
        len("".join(filter(str.isdigit, v))) for v in raw_ids_str if v
    ]
    id_width = max(digit_lengths) if digit_lengths else 5

    emp_blocks = {}
    current_emp_norm = None

    for col_idx, raw in enumerate(row_vals, start=1):
        if raw not in (None, "") and str(raw).strip() != "":
            # New employee start
            emp_norm = _norm_emp_id(raw, id_width)

            # Close previous employee block if applicable
            if current_emp_norm is not None:
                emp_blocks[current_emp_norm]["end_col"] = col_idx - 1

            # Start new block
            emp_blocks[emp_norm] = {
                "raw_id": raw,
                "start_col": col_idx,
                "end_col": None,  # to be filled later
            }
            current_emp_norm = emp_norm

    # Close the last employee block, if any
    if current_emp_norm is not None:
        if emp_blocks[current_emp_norm]["end_col"] is None:
            emp_blocks[current_emp_norm]["end_col"] = last_col

    return emp_blocks, id_width


def get_feuil1_rows_info(
    wb,
    field_col: int = DEFAULT_FIELD_COL,
    field_row_start: int = DEFAULT_FIELD_ROW_START,
):
    """
    Scan Feuil1 to build:
      - rows_info: list of {row_index, code, label_raw, label_norm}
      - code_to_rows: map code_str -> [row indices]
      - labelnorm_to_rows: map normalized label -> [row indices]

    This is used by the mapping engine and by GUIs that want to list
    codes/labels for user selection.
    """
    sh = wb.sheets[SRC_SHEET]

    last_row = max(
        sh.range("A" + str(sh.cells.rows.count)).end("up").row,
        sh.range("B" + str(sh.cells.rows.count)).end("up").row,
    )

    rows_info = []
    code_to_rows = {}
    labelnorm_to_rows = {}

    for r in range(1, last_row + 1):
        code = sh.range((r, 1)).value
        label_raw = sh.range((r, field_col)).value

        if not code and not label_raw:
            continue

        code_str = str(code).strip() if code else ""
        label_str = str(label_raw).strip() if label_raw else ""
        label_norm = _norm_label(label_str)

        rows_info.append(
            {
                "row_index": r,
                "code": code_str,
                "label_raw": label_str,
                "label_norm": label_norm,
            }
        )

        if code_str:
            code_to_rows.setdefault(code_str, []).append(r)
        if label_norm:
            labelnorm_to_rows.setdefault(label_norm, []).append(r)

    return rows_info, code_to_rows, labelnorm_to_rows


# ---------------------------------------------------------------------------
# 3. Label harmonization + Feuil1 record extraction
# ---------------------------------------------------------------------------

def get_label_map():
    """
    Central harmonization table: group label variants into canonical names.
    """
    return {
        # Basic salary fields
        "salaire brut total": "salaire brut",
        "salaire brut": "salaire brut",
        "salaire de base mensuel": "salaire de base",
        "salaire de base": "salaire de base",

        # Contributions (employee)
        "cot salarie": "cotisations salarie",
        "cotisations salarie": "cotisations salarie",
        "salarial": "cotisations salarie",
        "total salarial": "cotisations salarie",

        # Contributions (employer)
        "cot patronale": "cotisations patronales",
        "cotisations patronales": "cotisations patronales",
        "patronal": "cotisations patronales",
        "total patronal": "cotisations patronales",

        # Taxes / PAS
        "net a payer": "net paye",
        "net paye": "net paye",
        "net imposable": "net imposable",
        "pas": "pas",
        "prelevement a la source": "pas",
        "impot": "pas",

        # Benefits
        "avantage": "avantages",
        "avantages": "avantages",
        "avantages en nature": "avantages",
    }


def extract_feuil1_records(
    wb,
    sheet1_meta,
    label_map=None,
    id_row: int = DEFAULT_ID_ROW,
    field_col: int = DEFAULT_FIELD_COL,
    field_row_start: int = DEFAULT_FIELD_ROW_START,
):
    """
    Extract all field values from Feuil1 per employee, using *dynamic
    employee blocks*.

    For each employee:
      - Dynamically detect the block [start_col, end_col] from the ID row.
      - For each column in that block, read vertical field values under the
        label column.
      - Combine into a dict[field_label_normalized] = value (last non-empty
        wins if multiple columns supply a value).

    Also computes special fields:
      - cotisations salarie
      - cotisations patronales
      - pas (by summing row 75 across the block)
      - avantages (by summing rows 66–74 across the block)

    Returns:
      records: dict[emp_norm] -> {canonical_field_label: value}
      matched_labels: list of canonical labels that are present on both
                      Feuil1 and Sheet1.
    """
    if label_map is None:
        label_map = get_label_map()

    sh_src = wb.sheets[SRC_SHEET]

    tgt_headers_norm = sheet1_meta["headers_norm"]

    # --- Feuil1 field labels (column B from row 5 down) ---
    field_names_raw = sh_src.range((field_row_start, field_col)).expand("down").value
    if not isinstance(field_names_raw, list):
        field_names_raw = [field_names_raw]
    field_names_norm = [_norm_label(f) for f in field_names_raw if f]

    # Map Feuil1 labels to canonical labels
    field_names_mapped = [label_map.get(name, name) for name in field_names_norm]

    # Determine which fields are shared with Sheet1
    matched_labels = list(set(field_names_mapped) & set(tgt_headers_norm))

    # Ensure special calculated fields are included when present
    for special in ["cotisations salarie", "cotisations patronales", "pas", "avantages"]:
        if special in tgt_headers_norm and special not in matched_labels:
            matched_labels.append(special)

    # --- Employee blocks ---
    emp_blocks, id_width = detect_employee_blocks(wb, id_row=id_row)

    records = {}

    for emp_norm, info in emp_blocks.items():
        start_col = info["start_col"]
        end_col = info["end_col"]
        rec = {}

        # --- Read the standard vertical fields for each column in this block ---
        for col in range(start_col, end_col + 1):
            vals = sh_src.range(
                (field_row_start, col),
                (field_row_start + len(field_names_mapped) - 1, col),
            ).value

            if not isinstance(vals, list):
                vals = [vals]

            for field, val in zip(field_names_mapped, vals):
                if val not in (None, ""):
                    # last non-empty value across the block wins
                    rec[field] = val

        # --- Special computed fields (still using fixed rows, but dynamic block width) ---
        try:
            # Employee contributions (salarial): row 5, second column of block
            if start_col + 1 <= end_col:
                rec["cotisations salarie"] = sh_src.range((5, start_col + 1)).value

            # Employer contributions: row 5, third column of block
            if start_col + 2 <= end_col:
                rec["cotisations patronales"] = sh_src.range((5, start_col + 2)).value

            # PAS: sum row 75 across the entire employee block
            pas_vals = sh_src.range((75, start_col), (75, end_col)).value
            pas_total = 0
            if isinstance(pas_vals, list):
                for v in pas_vals:
                    if isinstance(v, (int, float)):
                        pas_total += v
            else:
                if isinstance(pas_vals, (int, float)):
                    pas_total += pas_vals
            rec["pas"] = pas_total

            # Avantages: sum rows 66–74 across the entire block
            avantage_vals = sh_src.range((66, start_col), (74, end_col)).value
            total_avantage = 0
            if isinstance(avantage_vals, list):
                # 2D or 1D list; handle both
                for row in avantage_vals:
                    if isinstance(row, list):
                        for v in row:
                            if isinstance(v, (int, float)):
                                total_avantage += v
                    else:
                        if isinstance(row, (int, float)):
                            total_avantage += row
            else:
                if isinstance(avantage_vals, (int, float)):
                    total_avantage += avantage_vals
            rec["avantages"] = total_avantage

        except Exception as e:
            print(f"⚠️ Warning while processing special fields for employee {emp_norm}: {e}")

        records[emp_norm] = rec

    return records, matched_labels


# ---------------------------------------------------------------------------
# 4. Combined employee layout: Feuil1 blocks + Sheet1 rows
# ---------------------------------------------------------------------------

def get_employee_layout_and_sheet1(
    wb,
    id_row: int = DEFAULT_ID_ROW,
    tgt_sheet_name: str = TGT_SHEET,
):
    """
    Build:
      - emp_to_feuil1: map emp_norm -> (start_col, end_col)
      - emp_to_sheet1: map emp_norm -> Sheet1 row index
      - header_map: from Sheet1 metadata

    This is used by the mapping engine to know:
      - which columns of Feuil1 to sum for an employee
      - which row in Sheet1 to write results to
    """
    meta = get_sheet1_meta(wb, tgt_sheet_name=tgt_sheet_name)
    emp_blocks, _ = detect_employee_blocks(wb, id_row=id_row)

    emp_to_feuil1 = {
        emp_norm: (info["start_col"], info["end_col"])
        for emp_norm, info in emp_blocks.items()
    }

    emp_to_sheet1 = {}
    for row_idx, emp_norm in enumerate(meta["ids_norm"], start=2):
        if emp_norm:
            emp_to_sheet1[emp_norm] = row_idx

    return emp_to_feuil1, emp_to_sheet1, meta["header_map"]


# ---------------------------------------------------------------------------
# 5. Analytics helper: field time series from Sheet1
# ---------------------------------------------------------------------------

def get_field_series_from_sheet1(
    wb,
    field_label: str,
    tgt_sheet_name: str = TGT_SHEET,
):
    """
    Return (employee_ids, values, average) for a given logical field label.

    Uses normalized labels for robustness. Any non-numeric values are treated
    as 0 for the purpose of computing stats/plots.
    """
    meta = get_sheet1_meta(wb, tgt_sheet_name=tgt_sheet_name)
    sh = meta["sheet"]

    target_norm = _norm_label(field_label)
    header_map = meta["header_map"]

    if target_norm not in header_map:
        raise ValueError(f"Field '{field_label}' not found in Sheet '{tgt_sheet_name}'")

    col_num = header_map[target_norm]

    emp_ids = meta["ids"]
    vals = sh.range((2, col_num)).expand("down").value
    if not isinstance(vals, list):
        vals = [vals]

    numeric_values = [v if isinstance(v, (int, float)) else 0 for v in vals]

    avg = sum(numeric_values) / len(numeric_values) if numeric_values else 0

    return emp_ids, numeric_values, avg
