import win32com.client as win32
import sys

try:
    excel = win32.GetActiveObject("Excel.Application")
    wb = excel.ActiveWorkbook
    print(f"Active workbook: {wb.Name}")
    
    # Check if workbook is protected
    if wb.ProtectStructure:
        print("Workbook structure IS protected. Unprotecting...")
        wb.Unprotect("Z961814r")
    else:
        print("Workbook structure is NOT protected.")
        
    found = False
    for ws in wb.Worksheets:
        if "הוראות" in ws.Name:
            print(f"Found sheet: {ws.Name}")
            ws.Tab.Color = 255 + (153 * 256) + (153 * 65536)  # BGR format for Excel COM: RGB(255,153,153) is 153*65536 + 153*256 + 255 = 10066431. Wait!
            # Let's use the exact integer value for RGB(255, 153, 153)
            # In VBA, RGB(R, G, B) = R + (G * 256) + (B * 65536)
            ws.Tab.Color = 255 + (153 * 256) + (153 * 65536)
            print("Tab color changed successfully.")
            found = True
            break
            
    if not found:
        print("Sheet not found.")
        
except Exception as e:
    print(f"Error: {e}")
