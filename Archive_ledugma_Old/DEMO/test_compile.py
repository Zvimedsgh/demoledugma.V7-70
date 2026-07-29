import win32com.client
import os

try:
    xlApp = win32com.client.DispatchEx('Excel.Application')
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    wb = xlApp.Workbooks.Add()
    xlmodule = wb.VBProject.VBComponents.Import(r'C:\ledugma\DEMO\modDemoReports_V9.40.bas')
    
    # Try to compile the VBA project
    # There is no direct "Compile" method in the VBProject object model that returns success/failure easily,
    # but if we run a dummy macro, it forces a compile. If there is a syntax error, it throws.
    # We can try to run A00_SetupMainSheet_V9_40 (it will fail at runtime because sheets don't exist, 
    # but a syntax error will throw a different COM error or before execution).
    try:
        xlApp.Run('A00_SetupMainSheet_V9_40')
        print("Run succeeded (unexpected, should fail on missing sheets).")
    except Exception as e:
        print("Run failed with:", str(e))
        
    wb.Close(False)
    xlApp.Quit()
except Exception as e:
    print("Script failed:", str(e))
