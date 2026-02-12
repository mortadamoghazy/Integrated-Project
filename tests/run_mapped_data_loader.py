"""
run_mapped_data_loader.py

Convenience script to convert mapped Excel data to CSV format
for plotting, forecasting, and analysis.

Usage:
    python scripts/run_mapped_data_loader.py
    python scripts/run_mapped_data_loader.py <input.xlsx> <output.csv>
"""

import sys
import os
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.core.mapped_data_loader import load_all_mapped_months


def main():
    # FIX_ME: Default input/output paths - Update if your files are located elsewhere
    # default_input: The sorted Excel file containing mapped payroll data
    # default_output: Where the consolidated CSV file will be saved
    # These can also be overridden via command line arguments:
    #   python run_mapped_data_loader.py <input.xlsx> <output.csv>
    default_input = project_root / "data" / "raw" / "data_sorted.xlsx"
    default_output = project_root / "outputs" / "payroll_mapped_long.csv"
    
    # Parse command line arguments
    if len(sys.argv) > 1:
        input_path = Path(sys.argv[1])
    else:
        input_path = default_input
    
    if len(sys.argv) > 2:
        output_path = Path(sys.argv[2])
    else:
        output_path = default_output
    
    # Check if input exists
    if not input_path.exists():
        print(f"❌ Input file not found: {input_path}")
        print()
        print("Expected location: data/raw/data_sorted.xlsx")
        print("This file is created by the mapping engine.")
        print()
        print("To create it:")
        print("  1. Run import_dialog.py to import and map your data")
        print("  2. Or run test_monthly_mapping_engine.py to apply mappings")
        return 1
    
    print("=" * 70)
    print(" MAPPED DATA LOADER - Convert Excel to CSV for Analysis")
    print("=" * 70)
    print()
    print(f"📂 Input:  {input_path}")
    print(f"💾 Output: {output_path}")
    print()
    print("=" * 70)
    print()
    
    try:
        df = load_all_mapped_months(str(input_path), str(output_path))
        
        print()
        print("=" * 70)
        print(" ✅ SUCCESS - Data Ready for Analysis")
        print("=" * 70)
        print()
        print(f"📊 Rows:      {len(df):,}")
        print(f"📋 Columns:   {len(df.columns)}")
        print(f"👥 Employees: {df['employee_id'].nunique()}")
        print(f"📅 Months:    {df['month'].nunique()}")
        print()
        print("📋 Columns available:")
        for col in df.columns:
            print(f"   • {col}")
        print()
        print("=" * 70)
        print()
        print("Next steps:")
        print("  • Use for plotting: python scripts/run_plot.py")
        print("  • Use for forecasting: python scripts/run_forecast.py")
        print(f"  • CSV location: {output_path}")
        print()
        
        return 0
        
    except Exception as e:
        print()
        print("=" * 70)
        print(" ❌ ERROR")
        print("=" * 70)
        print()
        print(f"Error: {e}")
        print()
        import traceback
        traceback.print_exc()
        print()
        return 1


if __name__ == "__main__":
    sys.exit(main())
