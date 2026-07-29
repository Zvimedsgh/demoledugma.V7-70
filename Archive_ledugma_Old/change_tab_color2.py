import win32com.client as win32
import time
import sys

success = False
for i in range(5):
    try:
        excel = win32.GetActiveObject("Excel.Application")
        wb = excel.ActiveWorkbook
        
        if wb is None:
            print("No active workbook found.")
            time.sleep(1)
            continue
            
        print(f"Active workbook: {wb.Name}")
        
        if wb.ProtectStructure:
            print("Unprotecting workbook...")
            wb.Unprotect("Z961814r")
            
        for ws in wb.Worksheets:
            if "הוראות" in ws.Name:
                ws.Unprotect("Z961814r")
                print(f"Found sheet: {ws.Name}")
                # RGB(255, 153, 153) in BGR integer:
                bgr = 255 + (153 * 256) + (153 * 65536)
                ws.Tab.Color = bgr
                print(f"Tab color successfully changed to BGR={bgr}.")
                success = True
                
        if wb.ProtectStructure == False:
            wb.Protect("Z961814r")
            print("Workbook reprotected.")
            
        if success:
            break
            
    except Exception as e:
        print(f"Attempt {i+1} Error: {e}")
        time.sleep(1)

if not success:
    print("Could not change the tab color after 5 attempts.")
