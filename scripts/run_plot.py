import sys
import os
import time

# Bootstrap sys.path to include project root
BASE_DIR = os.path.dirname(os.path.dirname(__file__))
sys.path.insert(0, BASE_DIR)

print("PYTHON ENTRY POINT: run_plot.py called")
print("BASE_DIR =", BASE_DIR)

try:
    from src.features.analytics.plotting_gui import main
except Exception as e:
    print("IMPORT ERROR:", e)
    time.sleep(20)
    raise

try:
    main()
except Exception as e:
    print("Runtime ERROR:", e)
    time.sleep(20)
