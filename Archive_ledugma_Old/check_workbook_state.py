import win32com.client as win32
import time
import sys

try:
    excel = win32.GetActiveObject("Excel.Application")
    wb = excel.ActiveWorkbook
    
    if wb is None:
        print("No active workbook found.")
        sys.exit()
        
    print(f"Active workbook: {wb.Name}")
    print(f"Workbook ProtectStructure: {wb.ProtectStructure}")
    print(f"Workbook ProtectWindows: {wb.ProtectWindows}")
    print(f"Workbook MultiUserEditing (Legacy Shared): {wb.MultiUserEditing}")
    print(f"Workbook ReadOnly: {wb.ReadOnly}")
    
    # Selected sheets count
    sel_count = excel.ActiveWindow.SelectedSheets.Count
    print(f"Number of selected sheets: {sel_count}")
    
    # Try to unprotect again just in case
    if wb.ProtectStructure:
        print("Attempting to unprotect workbook structure with Z961814r...")
        try:
            wb.Unprotect("Z961814r")
            print(f"After unprotect -> ProtectStructure: {wb.ProtectStructure}")
        except Exception as e:
            print(f"Failed to unprotect: {e}")
            
    # Now try to set the tab color of "הוראות"
    for ws in wb.Worksheets:
        if "הוראות" in ws.Name:
            print(f"Sheet '{ws.Name}' ProtectContents: {ws.ProtectContents}")
            try:
                bgr = 255 + (153 * 256) + (153 * 65536)
                ws.Tab.Color = bgr
                print(f"Successfully forced Tab Color to pastel red!")
            except Exception as e:
                print(f"Failed to set Tab Color: {e}")

except Exception as e:
    print(f"COM Error: {e}")
