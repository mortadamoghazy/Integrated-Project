"""
automation.py
Reads Feuil1, extracts payroll data, and writes to Sheet1 / Sheet2.
Uses src.core.data_loader for all low-level data extraction.
"""

import xlwings as xw

from src.core.config import SRC_SHEET, TGT_SHEET, PAYROLL_FILE
from src.core.data_loader import (
    get_sheet1_meta,
    extract_feuil1_records,
    get_label_map,
)
from src.core.normalization import _norm_label, _norm_emp_id


def fill_simplified_table():
    """
    Main routine:
    - Attach to workbook.
    - Use data_loader to read Sheet1 metadata & Feuil1 records.
    - Write extracted values to Sheet1 and Sheet2 based on matching labels.
    """
    try:
        wb = xw.Book.caller()
        print("✅ Attached via Book.caller()")
    except Exception:
        print("⚠️ Not called from Excel — attaching manually...")
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        try:
            wb = [b for b in app.books if b.name.lower() == "pay employier sept25.xlsm"][0]
        except IndexError:
            wb = app.books.open(PAYROLL_FILE)

    sh_tgt1 = wb.sheets[TGT_SHEET]

    print("Connected workbook:", wb.name)
    print("Full path:", wb.fullname)

    # --- Load Sheet1 metadata via data_loader ---
    sheet1_meta = get_sheet1_meta(wb, tgt_sheet_name=TGT_SHEET)

    # --- Use data_loader to extract all Feuil1 records with dynamic blocks ---
    label_map = get_label_map()
    records, matched_labels = extract_feuil1_records(
        wb,
        sheet1_meta,
        label_map=label_map,
    )

    tgt_ids = sheet1_meta["ids"]
    tgt_ids_norm = sheet1_meta["ids_norm"]

    print("\n========== Extracted Data Summary ==========")
    for emp_id, rec in records.items():
        print(f"\nEmployee {emp_id}:")
        for field, val in rec.items():
            if field in matched_labels:
                print(f"   {field}: {val}")
    print("============================================\n")

    def write_to_sheet(sheet):
        header_map = sheet1_meta["header_map"]

        for i, (emp_raw, emp_norm) in enumerate(zip(tgt_ids, tgt_ids_norm), start=2):
            if emp_norm not in records:
                continue

            rec = records[emp_norm]
            for field_name in matched_labels:
                col_num = header_map.get(field_name)
                if not col_num or col_num <= 1:
                    continue  # skip column A or missing headers

                value = rec.get(field_name, None)
                if value not in (None, ""):
                    cell = sheet.range(i, col_num)
                    cell.value = value
                    cell.color = (255, 255, 153)
                    cell.api.Font.Color = 0
                    cell.api.Font.Bold = True

    print("✏️ Writing data to Sheet1...")
    write_to_sheet(sh_tgt1)


    wb.save()
    wb.app.calculate()
    xw.apps.active.api.StatusBar = "✅ Data written successfully to Sheet1."
    print("✅ Done — data transferred and workbook saved.")


if __name__ == "__main__":
    import os
    import traceback

    try:
        print("🟢 Running in standalone mode...")
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        try:
            wb = [b for b in app.books if b.name.lower() == "pay employier sept25.xlsm"][0]
        except IndexError:
            wb = app.books.open(PAYROLL_FILE)

        wb.set_mock_caller()
        fill_simplified_table()

        print("✅ Completed successfully.")
        os.system("pause")
    except Exception:
        print("❌ An error occurred:\n")
        traceback.print_exc()
        os.system("pause")
