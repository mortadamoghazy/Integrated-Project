Attribute VB_Name = "PayrollControlPanel"
' =====================================================
' PAYROLL ANALYTICS CONTROL PANEL - VBA MODULE
' =====================================================
' This module provides buttons to:
' 1. Import and map payroll data from Excel files
' 2. Open the outputs folder
' 3. Open the mapped workbook
' 4. Launch the analytics dashboard
'
' INSTALLATION INSTRUCTIONS:
' 1. Open Control_Panel_New.xlsm in Excel
' 2. Press ALT+F11 to open VBA Editor
' 3. File > Import File > Select this .bas file
' 4. Update the FIX_ME paths below
' 5. Save and close VBA Editor
' =====================================================

' =====================================================
' FIX_ME: PYTHON AND PROJECT PATHS
' =====================================================
' UPDATE THESE TWO PATHS TO MATCH YOUR LOCAL ENVIRONMENT
'
' PYTHON_PATH: Full path to your Python executable
' How to find it:
'   1. Open Command Prompt or PowerShell
'   2. Run: python -c "import sys; print(sys.executable)"
'   3. Copy the output path here
'
' Common locations:
'   - Anaconda: C:\Users\YourName\Anaconda3\python.exe
'   - Standard: C:\Users\YourName\AppData\Local\Programs\Python\Python310\python.exe
'   - Virtual env: C:\Users\YourName\Documents\Project\venv\Scripts\python.exe
'
' PROJECT_PATH: Full path to your project root folder
' This is the "Integrated Project" folder containing:
'   - data/ folder
'   - outputs/ folder
'   - src/ folder
'   - scripts/ folder
'
' IMPORTANT: Must end with backslash (\)
' Example: "C:\Users\YourName\Documents\MARS\Integrated Project\"
' =====================================================

Const PYTHON_PATH As String = "C:\Users\Morta\AppData\Local\Programs\Python\Python310\python.exe"
Const PROJECT_PATH As String = "C:\Users\Morta\OneDrive\Desktop\Gam3a\MARS\Integrated Project\"

' =====================================================
' BUTTON 1: Import and Map Data
' CSV GENERATES AUTOMATICALLY - NO SEPARATE BUTTON NEEDED
' =====================================================
Sub ImportAndMapData()
    Dim scriptPath As String
    Dim cmd As String
    
    MsgBox "Launching Import & Mapping GUI..." & vbCrLf & vbCrLf & _
           "The Python window will open. Follow the GUI to:" & vbCrLf & _
           "1. Select your Excel file from data/raw" & vbCrLf & _
           "2. Define mappings" & vbCrLf & _
           "3. Close GUI to apply mappings" & vbCrLf & vbCrLf & _
           "CSV will be generated automatically!", vbInformation, "Import & Map"
    
    scriptPath = PROJECT_PATH & "src\excel_tools\import_dialog.py"
    cmd = Chr(34) & PYTHON_PATH & Chr(34) & " " & Chr(34) & scriptPath & Chr(34)
    
    Call Shell(cmd, vbNormalFocus)
End Sub

' =====================================================
' BUTTON 2: Open Output Folder
' =====================================================
Sub OpenOutputFolder()
    Dim folderPath As String
    folderPath = PROJECT_PATH & "outputs"
    
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Outputs folder not found!" & vbCrLf & vbCrLf & _
               "Expected location: " & folderPath, vbExclamation, "No Folder"
        Exit Sub
    End If
    
    Shell "explorer.exe " & Chr(34) & folderPath & Chr(34), vbNormalFocus
End Sub

' =====================================================
' BUTTON 3: Open Mapped Workbook
' =====================================================
Sub OpenMappedWorkbook()
    Dim wbPath As String
    wbPath = PROJECT_PATH & "data\raw\data_sorted.xlsx"
    
    If Dir(wbPath) = "" Then
        MsgBox "Mapped workbook not found!" & vbCrLf & vbCrLf & _
               "Please run 'Import & Map Data' first." & vbCrLf & vbCrLf & _
               "Expected location: " & wbPath, vbExclamation, "No File"
        Exit Sub
    End If
    
    Workbooks.Open wbPath
    MsgBox "Mapped workbook opened!", vbInformation, "Success"
End Sub

' =====================================================
' BUTTON 4: Visualize Data
' =====================================================
Sub VisualizeData()
    Dim scriptPath As String
    Dim cmd As String
    
    MsgBox "Launching Payroll Analytics Dashboard..." & vbCrLf & vbCrLf & _
           "Select from 10 visualization options:" & vbCrLf & _
           "• Total Cost Trend" & vbCrLf & _
           "• Forecasting (AR(2)+Ridge Model)" & vbCrLf & _
           "• Cost per Employee" & vbCrLf & _
           "• Employee Clustering" & vbCrLf & _
           "• Distribution Analysis" & vbCrLf & _
           "• And 5 more insights!", vbInformation, "Analytics"
    
    scriptPath = PROJECT_PATH & "src\features\analytics\dashboard_gui.py"
    cmd = Chr(34) & PYTHON_PATH & Chr(34) & " " & Chr(34) & scriptPath & Chr(34)
    
    Call Shell(cmd, vbNormalFocus)
End Sub

' =====================================================
' OPTIONAL: Verify Configuration
' =====================================================
' Run this macro first to verify your paths are correct
Sub VerifyConfiguration()
    Dim msg As String
    Dim pythonExists As Boolean
    Dim projectExists As Boolean
    
    ' Check Python path
    pythonExists = (Dir(PYTHON_PATH) <> "")
    
    ' Check project path
    projectExists = (Dir(PROJECT_PATH, vbDirectory) <> "")
    
    msg = "Configuration Check:" & vbCrLf & vbCrLf
    
    ' Python status
    If pythonExists Then
        msg = msg & "✓ Python found: " & PYTHON_PATH & vbCrLf
    Else
        msg = msg & "✗ Python NOT found: " & PYTHON_PATH & vbCrLf
        msg = msg & "  Please update PYTHON_PATH constant!" & vbCrLf
    End If
    
    msg = msg & vbCrLf
    
    ' Project status
    If projectExists Then
        msg = msg & "✓ Project folder found: " & PROJECT_PATH & vbCrLf
    Else
        msg = msg & "✗ Project folder NOT found: " & PROJECT_PATH & vbCrLf
        msg = msg & "  Please update PROJECT_PATH constant!" & vbCrLf
    End If
    
    msg = msg & vbCrLf
    
    ' Overall status
    If pythonExists And projectExists Then
        msg = msg & vbCrLf & "✓ Configuration is correct! You're ready to go."
        MsgBox msg, vbInformation, "Configuration OK"
    Else
        msg = msg & vbCrLf & "✗ Configuration has errors. Please fix the paths above."
        MsgBox msg, vbCritical, "Configuration Error"
    End If
End Sub
