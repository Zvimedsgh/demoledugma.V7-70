import os
import win32com.client as win32
import time

excel = None
try:
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    
    # Find recent .xlsm files in OneDrive
    found_files = []
    for root, dirs, files in os.walk(r'C:\Users\Zvi'):
        if 'OneDrive' in root:
            for f in files:
                if f.lower().endswith('.xlsm'):
                    full_path = os.path.join(root, f)
                    mod_time = os.path.getmtime(full_path)
                    found_files.append((mod_time, full_path))
                    
    found_files.sort(reverse=True)
    
    print("Checking the 3 most recent files in OneDrive for the missing sheet...")
    for mod_time, fpath in found_files[:3]:
        print(f"\nFile: {os.path.basename(fpath)}")
        try:
            wb = excel.Workbooks.Open(fpath, ReadOnly=True)
            for ws in wb.Worksheets:
                print(f"  - {ws.Name}")
            wb.Close(False)
        except Exception as e:
            print(f"  Failed to open: {e}")
            
finally:
    if excel:
        excel.Quit()
