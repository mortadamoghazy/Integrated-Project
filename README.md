Integrated Project/
│
├── scripts/
│   ├── run_excel.py         # Entry point for automation + mapping GUI
│   └── run_plot.py          # Entry point for field-value plotting GUI
│
├── src/
│   ├── __init__.py
│   │
│   ├── core/
│   │   ├── __init__.py
│   │   ├── config.py        # Central paths, sheet names, project constants
│   │   ├── normalization.py # For normalizing labels and employee IDs
│   │   └── data_loader.py   # ***Central dynamic data extractor***
│   │                         # Scans Feuil1 for labels & dynamic employee blocks
│   │                         # Scans Sheet1 for headers & IDs
│   │                         # Provides robust, reusable extraction API
│   │
│   ├── features/
│   │   ├── __init__.py
│   │   │
│   │   ├── payroll_processing/
│   │   │   ├── __init__.py
│   │   │   ├── automation.py       # Uses data_loader to extract Feuil1 records
│   │   │   │                       # Fills Sheet1 and Sheet2 with payroll data
│   │   │   └── mapping_engine.py   # Uses data_loader to read Feuil1 rows + employee blocks
│   │   │                            # Applies user-defined mappings
│   │   │
│   │   ├── gui/
│   │   │   ├── __init__.py
│   │   │   └── excel_gui_launcher.py
│   │   │        # Orchestrates:
│   │   │        # 1. fill_simplified_table()
│   │   │        # 2. apply_saved_mappings()
│   │   │        # 3. GUI for creating custom mappings
│   │   │
│   │   ├── analytics/
│   │   │   ├── __init__.py
│   │   │   ├── plotting_gui.py     # GUI for selecting field to plot
│   │   │   └── plot_salary_field.py# Uses data_loader to extract field series
│   │   │                            # Draws colored bar graph + average line
│
├── data/
│   ├── raw/
│   │   └── Pay employier sept25.xlsm  # Controlled input
│   └── processed/                      # Future processed outputs
│
├── docs/
│   └── project_description.pdf        # Fiche projet
│
├── tests/
│   ├── test_data_loader.py            # (optional future tests)
│   ├── test_mapping.py
│   └── test_automation.py
│
├── README.md
└── requirements.txt
