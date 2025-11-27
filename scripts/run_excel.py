import sys
import os
import time

# =====================================================
# BOOTSTRAP PYTHONPATH SO "src" IS IMPORTABLE
# =====================================================
BASE_DIR = os.path.dirname(os.path.dirname(__file__))  # project root
sys.path.insert(0, BASE_DIR)

print("PYTHON ENTRY POINT: run_excel.py called")
print("BASE_DIR =", BASE_DIR)
print("sys.path =", sys.path)

# =====================================================
# Now safe to import project modules
# =====================================================
try:
    from src.features.gui.excel_gui_launcher import main
except Exception as e:
    print("IMPORT ERROR:", e)
    time.sleep(20)
    raise

# =====================================================
# Run the application
# =====================================================
try:
    main()
except Exception as e:
    print("Runtime ERROR:", e)
    time.sleep(20)
else:
    print("GUI should be open now")
    time.sleep(20)
