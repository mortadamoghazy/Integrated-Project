"""
Central configuration for paths and shared settings.
"""

import os

# Base directory of the project (Integrated Project)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))

# Data directories
DATA_DIR = os.path.join(BASE_DIR, "data")
DATA_RAW_DIR = os.path.join(DATA_DIR, "raw")
DATA_PROCESSED_DIR = os.path.join(DATA_DIR, "processed")

# FIX_ME: Update the filename to match your actual payroll Excel file
# This should be the main payroll workbook in the data/raw/ directory
# Example: "Pay emploier sept25.xlsm" or "Control_Panel.xlsm"
PAYROLL_FILE = os.path.join(DATA_RAW_DIR, "Pay emploier sept25.xlsm")

# FIX_ME: Update sheet names to match your Excel workbook structure
# These should match the actual sheet names in your payroll Excel file
SRC_SHEET = "Feuil1"       # source sheet with full payroll data
TGT_SHEET = "Sheet1"       # simplified / target sheet for processed data
CUSTOM_SHEET = "CustomMap" # sheet for saved column mappings (will be created automatically)
