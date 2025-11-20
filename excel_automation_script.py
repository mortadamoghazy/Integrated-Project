"""
excel_automation_script.py
-------------------------------------------------
Feuil1 layout
-------------
• Row 3 : employee IDs (each covers multiple columns horizontally)
• Column B, row 5 ↓ : field names (Salaire de base, Salaire brut, etc.)
• Each employee block = 3 columns (can include Salarial / Patronal / etc.)

Sheet1 & Sheet2 layout
----------------------
• Column A : employee numbers (e.g., 00014)
• Columns B→ : target fields to fill
"""

import re
import unicodedata
import xlwings as xw


# ---------- Helper functions ----------

def _strip_accents(s):
    """Remove accents so labels can be compared reliably."""
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def _norm_label(s):
    """
    Clean and normalize header/field labels:
    - lowercase
    - remove accents
    - collapse whitespace
    - keep only simple characters
    """
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = _strip_accents(s)
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^\w\s./-]", "", s)
    return s


def _norm_emp_id(x, width):
    """Convert an employee ID into a zero-padded numeric string (e.g., 00014)."""
    digits = re.sub(r"\D", "", str(x))
    return digits.zfill(width) if digits else ""


# ---------- Main function ----------

def fill_simplified_table():
    """
    Main routine:
    - Connect to workbook
    - Read fields and employee blocks on Feuil1
    - Normalize labels
    - Extract values for each employee
    - Match fields with Sheet1 / Sheet2 headers
    - Write extracted data to both sheets
    """
    try:
        wb = xw.Book.caller()
        print("✅ Attached via Book.caller()")
    except Exception:
        # Fallback when run manually instead of from Excel
        print("⚠️ Not called from Excel — attaching manually...")
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        try:
            # Try locating the workbook already open
            wb = [b for b in app.books if b.name.lower() == "pay emploier sept25.xlsm"][0]
        except IndexError:
            # Otherwise open it from disk
            wb = app.books.open(r"C:\Users\DELL\Pay emploier sept25.xlsm")

    sh_src = wb.sheets["Feuil1"]
    sh_tgt1 = wb.sheets["Sheet1"]
 

    print("Connected workbook:", wb.name)
    print("Full path:", wb.fullname)

    # --- Read headers and employee IDs on Sheet1 ---
    tgt_headers = sh_tgt1.range("A1").expand("right").value
    tgt_headers_norm = [_norm_label(h) for h in tgt_headers]

    tgt_ids = sh_tgt1.range("A2").expand("down").value
    if not isinstance(tgt_ids, list):
        tgt_ids = [tgt_ids]
    tgt_ids = [str(v).strip() if v else "" for v in tgt_ids]

    # Determine width needed to zero-pad employee IDs uniformly
    id_width = max(len(re.sub(r"\D", "", v)) for v in tgt_ids if v) if any(tgt_ids) else 5
    tgt_ids_norm = [_norm_emp_id(v, id_width) for v in tgt_ids]

    # --- Feuil1 layout parameters ---
    field_row_start = 5
    field_col = 2
    max_cols = 120
    block_width = 3  # each employee span on Feuil1

    # --- Read list of field names in Feuil1 ---
    field_names = sh_src.range((field_row_start, field_col)).expand("down").value
    field_names = [_norm_label(f) for f in field_names if f]

    # --- Label harmonization table ---
    label_map = {
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

    # Normalize and harmonize labels from Feuil1
    field_names_mapped = [label_map.get(name, name) for name in field_names]

    # Figure out which fields exist both on Feuil1 and target sheets
    matched_labels = list(set(field_names_mapped) & set(tgt_headers_norm))
    print(f"✅ Matched labels ({len(matched_labels)}): {matched_labels}")

    # Ensure special calculated fields are included when present
    for special in ["cotisations salarie", "cotisations patronales", "pas", "avantages"]:
        if special in tgt_headers_norm and special not in matched_labels:
            matched_labels.append(special)
    print(f"✅ Updated matched labels ({len(matched_labels)}): {matched_labels}")

    # --- Read Feuil1 employee ID row (row 3) ---
    row_vals = sh_src.range((3, 1), (3, max_cols)).value
    if not isinstance(row_vals, list):
        row_vals = [row_vals]

    records = {}  # employee → dict of extracted fields

    for col_idx, emp in enumerate(row_vals, start=1):
        if not emp or str(emp).strip() == "":
            continue  # skip empty columns with no employee

        emp_norm = _norm_emp_id(emp, id_width)
        rec = {}

        # --- Read the standard vertical fields (3-column block) ---
        for subcol in range(block_width):
            read_col = col_idx + subcol
            vals = sh_src.range(
                (field_row_start, read_col),
                (field_row_start + len(field_names_mapped) - 1, read_col)
            ).value

            for field, val in zip(field_names_mapped, vals):
                if val not in [None, ""]:
                    rec[field] = val

        # --- Read special computed values on fixed rows ---
        try:
            # Employee contributions (salarial)
            rec["cotisations salarie"] = sh_src.range((5, col_idx + 1)).value

            # Employer contributions
            rec["cotisations patronales"] = sh_src.range((5, col_idx + 2)).value

            # PAS: sum of 3 columns on row 75
            pas_vals = sh_src.range((75, col_idx), (75, col_idx + 2)).value
            if isinstance(pas_vals, list):
                rec["pas"] = sum(v for v in pas_vals if isinstance(v, (int, float)))
            else:
                rec["pas"] = pas_vals if isinstance(pas_vals, (int, float)) else 0

            # Benefits: sum rows 66–74 across 3 columns
            avantage_vals = sh_src.range((66, col_idx), (74, col_idx + 2)).value
            total_avantage = 0
            for row in avantage_vals:
                for v in (row if isinstance(row, list) else [row]):
                    if isinstance(v, (int, float)):
                        total_avantage += v
            rec["avantages"] = total_avantage

        except Exception as e:
            # Any issues reading special ranges per employee
            print(f"⚠️ Warning while processing employee {emp_norm}: {e}")

        records[emp_norm] = rec

    # --- Print summary of extracted data (for debugging) ---
    print("\n========== Extracted Data Summary ==========")
    for emp_id, rec in records.items():
        print(f"\nEmployee {emp_id}:")
        for field, val in rec.items():
            if field in matched_labels:
                print(f"   {field}: {val}")
    print("============================================\n")

    # --- Function to write data to a target sheet ---
    def write_to_sheet(sheet):
        # Map normalized header → column number
        header_map = {col_name: idx + 1 for idx, col_name in enumerate(tgt_headers_norm)}

        # Loop through employees on the target sheet
        for i, (emp_raw, emp_norm) in enumerate(zip(tgt_ids, tgt_ids_norm), start=2):
            if emp_norm not in records:
                continue

            rec = records[emp_norm]
            for field_name in matched_labels:
                col_num = header_map.get(field_name)
                if not col_num or col_num <= 1:
                    continue  # skip column A or missing headers

                value = rec.get(field_name, None)
                if value not in [None, ""]:
                    cell = sheet.range(i, col_num)
                    cell.value = value
                    # Highlight written cells
                    cell.color = (255, 255, 153)
                    cell.api.Font.Color = 0
                    cell.api.Font.Bold = True

    # --- Write results to both target sheets ---
    print("✏️ Writing data to Sheet1...")
    write_to_sheet(sh_tgt1)



    wb.save()
    wb.app.calculate()
    xw.apps.active.api.StatusBar = "✅ Data written successfully to Sheet1 and Sheet2."
    print("✅ Done — data transferred and workbook saved.")


# ---------- Standalone mode (for manual testing) ----------

if __name__ == "__main__":
    import os, traceback
    try:
        print("🟢 Running in standalone mode...")
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        try:
            wb = [b for b in app.books if b.name.lower() == "pay emploier sept25.xlsm"][0]
        except IndexError:
            wb = app.books.open(r"C:\Users\DELL\Pay emploier sept25.xlsm")

        # Simulate Excel calling the function
        wb.set_mock_caller()
        fill_simplified_table()

        print("✅ Completed successfully.")
        os.system("pause")
    except Exception:
        print("❌ An error occurred:\n")
        traceback.print_exc()
        os.system("pause")
