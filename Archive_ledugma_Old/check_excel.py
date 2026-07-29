import win32com.client

try:
    xl = win32com.client.GetActiveObject("Excel.Application")
    wb = xl.ActiveWorkbook
    print(f"Active Workbook: {wb.Name}")
    print(f"ProtectStructure: {wb.ProtectStructure}")
    print(f"ProtectWindows: {wb.ProtectWindows}")
except Exception as e:
    print(f"Error: {e}")
