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

# Main payroll Excel file
PAYROLL_FILE = os.path.join(DATA_RAW_DIR, "Pay emploier sept25.xlsm")

# Excel sheet names
SRC_SHEET = "Feuil1"       # source sheet with full payroll
TGT_SHEET = "Sheet1"       # simplified / target sheet
CUSTOM_SHEET = "CustomMap" # sheet for saved mappings
