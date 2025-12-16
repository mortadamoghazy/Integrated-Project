import pandas as pd
import unicodedata
import re
from pathlib import Path
import xlwings as xw


# -----------------------------
# NORMALIZATION UTILITIES
# -----------------------------
def normalize_label(label: str) -> str:
    """Normalize column names: lowercase, remove accents, simplify."""
    if label is None:
        return ""
    s = str(label).strip().lower()

    # remove accents
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )

    # replace spaces by underscores
    s = re.sub(r"\s+", "_", s)

    # remove non-alphanumeric/underscore
    s = re.sub(r"[^\w_]", "", s)

    return s


def normalize_employee_id(x):
    """Ensure employee_id stays as zero-padded string."""
    x = str(x).strip()
    digits = re.sub(r"\D", "", x)
    if digits == "":
        return ""
    return digits.zfill(5)


# -----------------------------
# SINGLE SHEET LOADER
# -----------------------------
def load_month_sheet(sh):
    """
    Robust loader: detects the header row by scanning for typical payroll fields.
    Does NOT assume headers are in row 1.
    """

    raw = sh.range("A1").expand().value

    # Normalize everything into list-of-lists
    if not isinstance(raw, list):
        raw = [raw]
    if raw and not isinstance(raw[0], list):
        raw = [raw]

    # keywords we expect in header row
    header_keywords = {
        "employee_id", "role", "salaire_brut", "cot_salariale",
        "cot_patronale", "net_imposable", "pas",
        "net_paye", "avantages", "total_cost"
    }

    header_row_idx = None

    for idx, row in enumerate(raw):
        cleaned = {normalize_label(str(x)) for x in row if x}
        if cleaned & header_keywords:
            header_row_idx = idx + 1  # Excel is 1-based
            break

    if header_row_idx is None:
        print("❌ Could not find header row in sheet:", sh.name)
        return pd.DataFrame()

    # Load dataframe beginning from detected header row
    df = sh.range((header_row_idx, 1)).expand().options(
        pd.DataFrame,
        header=True,
        index=False
    ).value

    # Normalize column names
    df.columns = [normalize_label(c) for c in df.columns]

    # Add month (sheet name, e.g. "2023-01")
    df["month"] = sh.name

    # Normalize employee IDs
    if "employee_id" in df.columns:
        df["employee_id"] = (
            df["employee_id"]
            .astype(str)
            .str.strip()
            .apply(normalize_employee_id)
        )

    return df


# -----------------------------
# FULL WORKBOOK LOADER
# -----------------------------
def load_all_months(workbook_path):
    """
    Loads all sheets from a multi-month payroll workbook.
    - Dynamically detects columns
    - Normalizes headers
    - Adds month from sheet name
    - Concatenates all sheets
    """

    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook {workbook_path} not found")

    app = xw.App(visible=False)
    wb = app.books.open(str(workbook_path))

    all_dataframes = []

    try:
        for sh in wb.sheets:
            sh_name = sh.name

            if sh_name.lower() in {"summary", "readme", "info"}:
                continue

            print(f"🔎 Loading sheet: {sh_name}")

            try:
                df_month = load_month_sheet(sh)
                if df_month.empty:
                    print(f"⚠️ Sheet {sh_name} is empty – skipped.")
                    continue

                all_dataframes.append(df_month)
                print(f"   Loaded {len(df_month)} rows")

            except Exception as e:
                print(f"❌ Error in sheet {sh_name}: {e}")

    finally:
        wb.close()
        app.quit()

    if not all_dataframes:
        raise RuntimeError("No sheets were loaded – workbook has no valid data.")

    df = pd.concat(all_dataframes, ignore_index=True)

    print(f"📊 Total rows loaded: {len(df)}")

    # ---------------------------------------------------------
    # MONTH HANDLING FOR REGRESSION + NICE CSV FORMAT
    # ---------------------------------------------------------

    # 1) Convert month from sheet name (e.g. "2023-01") to datetime
    #    for correct chronological sorting and later regression.
    df["month"] = pd.to_datetime(df["month"], format="%Y-%m", errors="coerce")

    # 2) Sort by employee + month
    df = df.sort_values(["employee_id", "month"]).reset_index(drop=True)

    # 3) Convert month back to "YYYY-MM" string for the CSV
    #    (this is just for nicer display; regression code will convert to datetime again)
    df["month"] = df["month"].dt.strftime("%Y-%m")

    return df


# -----------------------------
# Example usage
# -----------------------------
if __name__ == "__main__":
    path = "YOUR_WORKBOOK.xlsx"  # Replace with your actual file
    df = load_all_months(path)

    # Save final consolidated CSV
    df.to_csv("payroll_long.csv", index=False)
    print("✅ Saved: payroll_long.csv")
