# Integrated-Project
Integrated Project – Excel Automation Tool

This project is part of the ENS3 MARS / ASI Integrated Project, supervised by Bogdan Robu (GIPSA-lab).
The goal is to build an interactive tool to help NGOs and small organizations manage:

Salary costs

Contributions

Project involvement

Monthly/annual budgeting

Scenario simulation

Reporting and forecasting

This repository contains the first complete working block of the project:
💡 Extract payroll data from Feuil1 and automatically populate structured tables in Sheet1 from Excel using Python.

_____________________________________________________________________________________________________

1. Full Excel → Python automation

A button inside Excel (Sheet2) launches a Python script using VBA.
This script:

Opens the workbook

Reads raw payroll data from Feuil1

Extracts salaries, contributions, benefits, and PAS

Normalizes labels

Matches fields to Sheet1

Fills Sheet1 automatically

Highlights all filled cells (yellow + bold)

This pipeline is now fully working.

_____________________________________________________________________________________________________

2.Sheet2 filling disabled

Data is now written only to Sheet1, by request.

_____________________________________________________________________________________________________
3. No GUI yet