import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

filepath = r"C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70.xlsm"
try:
    wb = excel.Workbooks.Open(filepath)
    print("Opened workbook")
    
    # Iterate through VBA components
    for comp in wb.VBProject.VBComponents:
        if comp.Type == 1: # Standard Module
            lines = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            if '"F20"' in lines:
                new_lines = lines.replace('"F20"', '"F16"')
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(new_lines)
                print(f"Updated module: {comp.Name}")
    
    wb.Save()
    wb.Close()
    print("Workbook saved and closed.")
except Exception as e:
    print(f"Error: {e}")
finally:
    excel.Quit()
