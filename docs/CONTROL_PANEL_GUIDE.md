# Excel Control Panel User Guide

## What Was Created

I've created **Control_Panel_New.xlsm** - a safe, VBA-based Excel interface with 4 buttons to automate your payroll mapping workflow.

### Why This Approach is Safe

**NO xlwings corruption risk!** Unlike the previous approach that caused system issues, this uses:
- **VBA Shell command** - launches Python as a separate process
- **No COM automation** - no shared memory between Excel and Python
- **Process isolation** - Python runs independently, Excel stays clean
- **No DLL hooks** - no low-level system integration

## The 4 Buttons

### Button 1: Import & Map Data
**What it does:**
1. Opens a file dialog to select your source Excel file
2. Launches the mapping GUI (tkinter interface)
3. Lets you define formulas using codes/labels and operators (+, -, *, /)
4. Automatically applies mappings when you close the GUI
5. Saves mappings to CustomMap sheet for next time

**When to use:** Every time you get a new payroll file or want to modify existing mappings

### Button 2: Generate CSV Report
**What it does:**
1. Reads the mapped data from data_sorted.xlsx
2. Converts all month sheets to a single CSV file
3. Formats data for analysis (adds role column, fixes month format, reorders columns)
4. Saves to outputs/payroll_mapped_long.csv

**When to use:** After mapping is complete, before running analysis/forecasting scripts

### Button 3: Open Output Folder
**What it does:**
1. Opens the outputs/ folder in Windows Explorer
2. Lets you view generated CSV files and analysis results

**When to use:** To quickly access your results without navigating folders

### Button 4: View Mapped Data
**What it does:**
1. Opens data_sorted.xlsx (the file with applied mappings)
2. Lets you inspect the calculated values in each month sheet

**When to use:** To verify mappings were applied correctly before generating CSV

## How to Use

### First Time Setup

1. **Open the file:**
   - Navigate to your project folder
   - Double-click `Control_Panel_New.xlsm`

2. **Enable macros:**
   - Excel will show a security warning at the top
   - Click "Enable Content" or "Enable Macros"
   - This is safe - it's your own VBA code

3. **Test the buttons:**
   - Start with Button 1 to import and map data
   - Then Button 2 to generate CSV
   - Use Buttons 3 and 4 to view results

### Normal Workflow

```
Button 1 (Import & Map) 
    ↓
[GUI opens, you define mappings, close GUI]
    ↓
[Mappings automatically apply to all months]
    ↓
Button 2 (Generate CSV)
    ↓
[CSV created in outputs/]
    ↓
Button 3 (View Outputs) or Button 4 (View Excel)
```

## Technical Details

### What Happens Behind the Scenes

**Button 1 execution:**
```vba
Shell "C:\...\python.exe" "C:\...\import_dialog.py"
```
- Python runs in its own window
- You interact with the GUI
- When you close GUI, Python applies mappings and exits
- Excel never touches Python's memory

**Button 2 execution:**
```vba
Shell "C:\...\python.exe" "C:\...\run_mapped_data_loader.py"
```
- Python reads data_sorted.xlsx
- Converts to CSV format
- Exits cleanly
- Excel stays isolated

### File Paths in VBA Code

The VBA code has two hardcoded paths:

```vba
Const PYTHON_PATH = "C:\Users\Morta\AppData\Local\Programs\Python\Python310\python.exe"
Const PROJECT_PATH = "C:\Users\Morta\OneDrive\Desktop\Gam3a\MARS\Integrated Project\"
```

**If these paths change**, you need to update them:
1. Open Control_Panel_New.xlsm
2. Press Alt+F11 (opens VBA editor)
3. Double-click Module1 in the Project Explorer
4. Edit the two Const lines at the top
5. Save and close

## Troubleshooting

### "Cannot find Python executable"
- Check PYTHON_PATH constant in VBA code
- Verify Python is installed at that location
- Update path if needed

### "Script not found"
- Check PROJECT_PATH constant in VBA code
- Make sure all scripts are in correct locations:
  - src/excel_tools/import_dialog.py
  - tests/run_mapped_data_loader.py

### "No mapped data found"
- You need to run Button 1 first (Import & Map)
- Make sure data_sorted.xlsx exists in data/raw/
- Check that mappings actually applied (Button 4 to view)

### GUI doesn't open
- Make sure tkinter is installed (comes with Python)
- Check for Python errors in the command window
- Verify import_dialog.py exists

### Macros disabled
- Open Excel Options → Trust Center → Trust Center Settings
- Go to Macro Settings
- Select "Enable all macros" or "Disable with notification"
- Close and reopen the file

## Safety Comparison

### Old Approach (xlwings) ❌
- Used COM automation
- Shared memory between Excel and Python
- Created DLL hooks
- Caused system corruption
- Required xlwings installation

### New Approach (Shell + VBA) ✅
- Uses simple Shell command
- Separate processes
- No shared memory
- No system hooks
- Only needs standard Python

## What You Can Customize

### Add More Buttons
1. Open VBA editor (Alt+F11)
2. Copy existing Sub (like ImportAndMapData)
3. Modify to call different Python script
4. Save, then insert new button on sheet
5. Assign your new Sub to the button

### Change Button Colors/Positions
1. Right-click any button
2. Format Control → Colors, Font, etc.
3. Move buttons by dragging

### Add More Instructions
- The sheet is editable
- Add rows, change text, add images
- Buttons will stay functional

## Summary

This Excel Control Panel provides a **safe, easy-to-use interface** for your payroll mapping workflow without any risk of system corruption. Each button launches Python scripts as separate processes, keeping Excel and Python completely isolated.

The workflow mirrors your previous manual process but automates everything through button clicks!
