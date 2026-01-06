"""
Quick launcher for payroll visualization dashboard
"""
import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.features.analytics.dashboard_gui import main

if __name__ == "__main__":
    print("🚀 Launching Payroll Analytics Dashboard...")
    main()
