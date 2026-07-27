import win32com.client
import sys

try:
    excel = win32com.client.GetActiveObject("Excel.Application")
    wb = excel.ActiveWorkbook
    
    print(f"Active Workbook: {wb.Name}")
    for ws in wb.Worksheets:
        visible_state = "Visible" if ws.Visible == -1 else ("Hidden" if ws.Visible == 0 else "VeryHidden")
        print(f"- {ws.Name} ({visible_state})")
        
        # Also check shapes on the 'הגדרות' sheet
        if ws.Name == chr(1492) + chr(1490) + chr(1491) + chr(1512) + chr(1493) + chr(1514):
            print("  Shapes on this sheet:")
            for shp in ws.Shapes:
                try:
                    print(f"    Shape: {shp.Name}, OnAction: {shp.OnAction}")
                except:
                    print(f"    Shape: {shp.Name}, OnAction: <error>")
except Exception as e:
    print(f"Error connecting to Excel: {e}")
