"""
mapped_data_loader.py

Loads mapped payroll data from the output workbook (created by the mapping engine)
and converts it to CSV format for plotting, forecasting, and analysis.

This is parallel to multi_sheet_loader.py but works with the new mapped data structure.
"""

import pandas as pd
import unicodedata
import re
from pathlib import Path
import xlwings as xw
from typing import Optional


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
def load_mapped_sheet(sh, month_name: str) -> pd.DataFrame:
    """
    Load a single mapped sheet. Assumes:
    - Row 1 contains headers (standardized by mapping engine)
    - Data starts from row 2
    - First column is employee_id
    """
    
    try:
        # Read the entire sheet starting from A1
        df = sh.range("A1").expand().options(
            pd.DataFrame,
            header=True,
            index=False
        ).value
        
        if df is None or df.empty:
            print(f"⚠️ Sheet {sh.name} is empty – skipped.")
            return pd.DataFrame()
        
        # Normalize column names
        df.columns = [normalize_label(c) for c in df.columns]
        
        # Add role column (real data doesn't have role info, so leave empty)
        if "role" not in df.columns:
            df.insert(1, "role", "")  # Insert after employee_id (position 1)
        
        # Add month column (use the sheet name as month identifier)
        df["month"] = month_name
        
        # Normalize employee IDs
        if "employee_id" in df.columns:
            df["employee_id"] = (
                df["employee_id"]
                .astype(str)
                .str.strip()
                .apply(normalize_employee_id)
            )
        
        # Remove rows where employee_id is empty
        if "employee_id" in df.columns:
            df = df[df["employee_id"] != ""]
        
        return df
        
    except Exception as e:
        print(f"❌ Error loading sheet {sh.name}: {e}")
        return pd.DataFrame()


# -----------------------------
# FULL WORKBOOK LOADER
# -----------------------------
def load_all_mapped_months(
    workbook_path: str,
    output_csv: str = "outputs/payroll_mapped_long.csv",
    exclude_sheets: Optional[list] = None
) -> pd.DataFrame:
    """
    Loads all mapped sheets from the output workbook.
    
    Args:
        workbook_path: Path to the mapped Excel workbook
        output_csv: Path where to save the consolidated CSV
        exclude_sheets: List of sheet names to exclude (e.g., ["CustomMap", "Summary"])
    
    Returns:
        Consolidated DataFrame with all months
    """
    
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook {workbook_path} not found")
    
    if exclude_sheets is None:
        exclude_sheets = ["custommap", "summary", "readme", "info", "template"]
    else:
        exclude_sheets = [s.lower() for s in exclude_sheets]
    
    print(f"📂 Opening workbook: {workbook_path}")
    
    app = xw.App(visible=False)
    wb = app.books.open(str(workbook_path.absolute()))
    
    all_dataframes = []
    
    try:
        for sh in wb.sheets:
            sh_name = sh.name
            
            # Skip excluded sheets
            if sh_name.lower() in exclude_sheets:
                print(f"⏭️ Skipping sheet: {sh_name}")
                continue
            
            print(f"🔎 Loading mapped sheet: {sh_name}")
            
            df_month = load_mapped_sheet(sh, month_name=sh_name)
            if df_month.empty:
                continue
            
            all_dataframes.append(df_month)
            print(f"   ✅ Loaded {len(df_month)} rows")
    
    finally:
        wb.close()
        app.quit()
    
    if not all_dataframes:
        raise RuntimeError("No sheets were loaded – workbook has no valid data.")
    
    df = pd.concat(all_dataframes, ignore_index=True)
    
    print(f"\n📊 Total rows loaded: {len(df)}")
    print(f"📅 Months found: {df['month'].nunique()}")
    print(f"👥 Employees found: {df['employee_id'].nunique()}")
    
    # ---------------------------------------------------------
    # MONTH HANDLING FOR REGRESSION + NICE CSV FORMAT
    # ---------------------------------------------------------
    
    # Store original month values (before any datetime conversion attempts)
    original_months = df["month"].copy()
    
    # Try to convert month to datetime
    # Support both "YYYY-MM" and "MM" formats
    try:
        # First try YYYY-MM format
        df["month"] = pd.to_datetime(df["month"], format="%Y-%m", errors="coerce")
        
        # If many are null, the format is likely just MM (month numbers)
        if df["month"].isnull().sum() > len(df) * 0.1:
            # For MM format, convert to YYYY-MM with base year 2025
            print("⚠️ Month format appears to be MM only. Converting to YYYY-MM with base year 2025.")
            # Convert to zero-padded strings and add year
            month_strings = original_months.astype(str).str.zfill(2)
            df["month"] = "2025-" + month_strings
            # Sort by employee + month
            df = df.sort_values(["employee_id", "month"]).reset_index(drop=True)
        else:
            # Successfully converted to datetime, sort and convert back to string
            df = df.sort_values(["employee_id", "month"]).reset_index(drop=True)
            # Convert month back to "YYYY-MM" string for the CSV
            df["month"] = df["month"].dt.strftime("%Y-%m")
    except Exception as e:
        print(f"⚠️ Could not parse month as datetime: {e}")
        print("   Converting to YYYY-MM format with base year 2025")
        # Fallback: use original values with year prefix
        month_strings = original_months.astype(str).str.zfill(2)
        df["month"] = "2025-" + month_strings
        df = df.sort_values(["employee_id", "month"]).reset_index(drop=True)
    
    # Reorder columns to match synthetic data format
    # Standard order: employee_id, role, salaire_brut, cot_salariale, cot_patronale, 
    #                 net_imposable, pas, net_paye, avantages, total_cost, month
    standard_cols = [
        "employee_id", "role", "salaire_brut", "cot_salariale", "cot_patronale",
        "net_imposable", "pas", "net_paye", "avantages", "total_cost", "month"
    ]
    
    # Collect columns that exist in the dataframe
    ordered_cols = [col for col in standard_cols if col in df.columns]
    
    # Add any custom columns (not in standard list) at the end
    custom_cols = [col for col in df.columns if col not in standard_cols]
    final_cols = ordered_cols + custom_cols
    
    df = df[final_cols]
    
    # Save CSV
    output_path = Path(output_csv)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False)
    print(f"\n✅ Saved: {output_path}")
    
    return df


# -----------------------------
# CLI Interface
# -----------------------------
if __name__ == "__main__":
    import sys
    
    # Default paths
    default_input = "data/raw/data_sorted.xlsx"
    default_output = "outputs/payroll_mapped_long.csv"
    
    if len(sys.argv) > 1:
        input_path = sys.argv[1]
    else:
        input_path = default_input
    
    if len(sys.argv) > 2:
        output_path = sys.argv[2]
    else:
        output_path = default_output
    
    print("=" * 60)
    print("MAPPED DATA LOADER - Excel to CSV Converter")
    print("=" * 60)
    print(f"Input:  {input_path}")
    print(f"Output: {output_path}")
    print("=" * 60)
    print()
    
    try:
        df = load_all_mapped_months(input_path, output_path)
        print()
        print("=" * 60)
        print("SUMMARY")
        print("=" * 60)
        print(f"Rows:      {len(df)}")
        print(f"Columns:   {len(df.columns)}")
        print(f"Employees: {df['employee_id'].nunique()}")
        print(f"Months:    {df['month'].nunique()}")
        print()
        print("Columns:", ", ".join(df.columns.tolist()))
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        sys.exit(1)
