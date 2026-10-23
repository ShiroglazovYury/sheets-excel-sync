Attribute VB_Name = "SheetSync"
' Buttons for sheets-excel-sync: Google Sheets -> Excel.
'
' Import this module into PERSONAL.XLSB (Alt+F11 -> File -> Import File).
' Requirements (see README, step 5):
'   * xlwings add-in installed and xlwings.bas imported into the same VBA project;
'   * xlwings ribbon -> Interpreter = <project>\.venv\Scripts\python.exe
'   * xlwings ribbon -> PYTHONPATH  = <project folder>
'
' The project path is NOT hardcoded here on purpose: it comes from the
' PYTHONPATH field, so moving the folder does not require editing VBA.

Sub SyncAll()
    RunPython "import sheetsync.excel_entry as e; e.run_all()"
End Sub

Sub SyncForecast()
    RunPython "import sheetsync.excel_entry as e; e.run('forecast')"
End Sub

Sub SyncOrders()
    RunPython "import sheetsync.excel_entry as e; e.run('orders')"
End Sub

' Write even if nothing changed or the row count dropped sharply.
Sub ForceSyncAll()
    RunPython "import sheetsync.excel_entry as e; e.run_all(force=True)"
End Sub
