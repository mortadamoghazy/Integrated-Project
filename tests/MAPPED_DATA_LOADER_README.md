# Mapped Data Loader

Converts mapped Excel payroll data to CSV format for plotting, forecasting, and analysis.

## Overview

This is the **parallel version** of `multi_sheet_loader.py` but for the **new mapped data workflow**:

- **Original workflow**: `multi_sheet_loader.py` → loads raw payroll data → `outputs/payroll_long.csv`
- **New workflow**: `mapped_data_loader.py` → loads mapped data → `outputs/payroll_mapped_long.csv`

Both workflows remain available for different use cases.

## Files

- **`src/core/mapped_data_loader.py`** - Core loader module
- **`scripts/run_mapped_data_loader.py`** - Convenience script with CLI

## Quick Start

### 1. Create Mapped Data

First, run the mapping process:

```bash
# Option A: Using import dialog (interactive)
python src/excel_tools/import_dialog.py

# Option B: Using test script (automatic)
python tests/test_monthly_mapping_engine.py
```

This creates: `data/raw/data_sorted.xlsx` (mapped workbook)

### 2. Convert to CSV

```bash
# Default paths
python scripts/run_mapped_data_loader.py

# Custom paths
python scripts/run_mapped_data_loader.py <input.xlsx> <output.csv>
```

Output: `outputs/payroll_mapped_long.csv`

## What It Does

1. **Opens** the mapped workbook (e.g., `data_sorted.xlsx`)
2. **Loads** all month sheets (excluding CustomMap, Summary, etc.)
3. **Normalizes** column names (lowercase, no accents, underscores)
4. **Consolidates** all months into one long-format DataFrame
5. **Saves** as CSV in `outputs/` directory

## Data Format

### Input (Excel)
```
Sheet "01":
employee_id | salaire_brut | cot_salariale | ...
00001       | 5000.00      | 500.00        | ...
00002       | 6000.00      | 600.00        | ...

Sheet "02":
employee_id | salaire_brut | cot_salariale | ...
00001       | 5100.00      | 510.00        | ...
00002       | 6100.00      | 610.00        | ...
```

### Output (CSV)
```
employee_id,salaire_brut,cot_salariale,...,month
00001,5000.00,500.00,...,01
00002,6000.00,600.00,...,01
00001,5100.00,510.00,...,02
00002,6100.00,610.00,...,02
```

## Usage in Analysis

The generated CSV can be used with existing analysis scripts:

```python
# Load the mapped data
import pandas as pd
df = pd.read_csv("outputs/payroll_mapped_long.csv")

# Use with regression models
from src.features.analytics.regression_ar1 import run_ar1_regression
results = run_ar1_regression(df)

# Use with plotting
from src.features.analytics.plot_salary_field import plot_employee_field
plot_employee_field(df, employee_id="00001", field="salaire_brut")
```

## Features

✅ **Supports custom columns** - Works with dynamically added labels  
✅ **Handles multiple month formats** - YYYY-MM or MM  
✅ **Normalizes employee IDs** - Zero-padded 5-digit strings  
✅ **Excludes system sheets** - Skips CustomMap, Summary, etc.  
✅ **Error handling** - Gracefully handles missing/empty sheets  
✅ **Progress feedback** - Shows what's being loaded  

## Configuration

Edit defaults in `run_mapped_data_loader.py`:

```python
default_input = "data/raw/data_sorted.xlsx"
default_output = "outputs/payroll_mapped_long.csv"
```

Or in `mapped_data_loader.py`:

```python
exclude_sheets = ["custommap", "summary", "readme", "info"]
```

## Comparison with Original Loader

| Feature | multi_sheet_loader.py | mapped_data_loader.py |
|---------|----------------------|----------------------|
| Input | Raw payroll Excel | Mapped payroll Excel |
| Output | payroll_long.csv | payroll_mapped_long.csv |
| Structure | Detects headers dynamically | Uses standardized headers |
| Columns | Variable per sheet | Consistent across sheets |
| Use Case | Original workflow | New mapping workflow |

## Troubleshooting

**File not found:**
```
❌ Input file not found: data/raw/data_sorted.xlsx
```
→ Run the mapping engine first to create the mapped workbook

**No sheets loaded:**
```
❌ No sheets were loaded – workbook has no valid data.
```
→ Check that month sheets exist and are not empty

**Column mismatch:**
- Mapped data should have consistent columns across sheets
- If custom columns were added, they'll be included in the CSV

## Next Steps

After generating the CSV:

1. **Plotting**: `python scripts/run_plot.py`
2. **Forecasting**: `python scripts/run_forecast.py`
3. **Analysis**: Use with regression models in `src/features/analytics/`

---

**Note**: Both workflows (original and new) coexist. Use the appropriate loader for your data source.
