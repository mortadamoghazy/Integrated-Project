import sys
from pathlib import Path

# Add project root to Python path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.append(str(PROJECT_ROOT))

from pathlib import Path
from src.core.multi_sheet_loader import load_all_months

def main():
    # FIX_ME: Update this path to point to your multi-month payroll workbook
    # This should be an Excel file (.xlsx or .xlsm) containing monthly payroll sheets
    # Default location: data/raw/synthetic_payroll_startup.xlsx
    wb_path = Path("data/raw/synthetic_payroll_startup.xlsx")

    print("=== STEP 1: Loading all months from workbook ===")
    df = load_all_months(str(wb_path))

    # Ensure outputs folder exists
    out_dir = Path("outputs")
    out_dir.mkdir(parents=True, exist_ok=True)

    # Save consolidated dataset
    excel_path = out_dir / "payroll_clean.xlsx"
    csv_path = out_dir / "payroll_long.csv"

    df.to_excel(excel_path, index=False)
    df.to_csv(csv_path, index=False)

    print("✅ Data loading completed.")
    print(f"   Saved Excel: {excel_path}")
    print(f"   Saved CSV:   {csv_path}")

if __name__ == "__main__":
    main()
