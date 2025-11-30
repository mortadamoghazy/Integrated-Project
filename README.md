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
_______________________________________________________________________________________________________________________________
Excel Button (Plot Field Values)
      │
      ▼
VBA: PlotButtonHandler()
      │
      ▼
python scripts/run_plot.py
      │
      ▼
analytics.plotting_gui.main()
      │
      ▼
User selects field
      │
      ▼
plot_salary_field.plot_field()
      ├── data_loader.get_field_series_from_sheet1()
      ├── matplotlib bar graph
      ├── average line
      ├── color coding: red/blue
      └── annotation of +/– differences
______________________________________________________________________________________________________________________________
Excel Button (Run Processing)
      │
      ▼
VBA: RunProcessButtonHandler()
      │
      ▼
python scripts/run_excel.py
      │
      ▼
excel_gui_launcher.main()
      ├── fill_simplified_table()
      │       ├── data_loader.get_sheet1_meta()
      │       ├── data_loader.extract_feuil1_records()
      │       └── writes data to Sheet1 + Sheet2
      │
      ├── apply_saved_mappings()
      │       ├── mapping_engine.load_saved_mappings()
      │       ├── data_loader.get_feuil1_rows_info()
      │       ├── data_loader.get_employee_layout_and_sheet1()
      │       └── mapping_engine.apply_single_mapping()
      │
      └── show_mapping_gui()
              ├── Lets user define a mapping
              ├── mapping_engine.apply_single_mapping()
              └── mapping_engine.save_or_update_mapping_row()
