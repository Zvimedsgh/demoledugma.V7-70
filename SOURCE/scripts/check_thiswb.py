import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

# Open the workbook
wb_path = r"c:\LEVAV PROJECT\Demo_Reports_System_V7.70 (2) (3).xlsm"
try:
    wb = excel.Workbooks.Open(wb_path)
    
    # Get the ThisWorkbook module
    for comp in wb.VBProject.VBComponents:
        if comp.Type == 100 and comp.Name == "ThisWorkbook":  # 100 = Document Module
            code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            print("--- ThisWorkbook ---")
            print(code)
            
    wb.Close(False)
except Exception as e:
    print(f"Error: {e}")
finally:
    excel.Quit()
