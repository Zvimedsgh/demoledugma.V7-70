import os
import win32com.client

base = r'C:\Users\Zvi'
target_dir = None
for f in os.listdir(base):
    if 'OneDrive' in f:
        od_path = os.path.join(base, f)
        if os.path.isdir(od_path):
            for sub in os.listdir(od_path):
                if 'דיווח' in sub:
                    target_dir = os.path.join(od_path, sub)

if target_dir:
    bait_file = os.path.join(target_dir, "Demo_Reports_Syatem_V7.70  (2)3.xlsm")
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(bait_file)
        ws = wb.Sheets(1)
        
        # Clear the warning text in B4
        ws.Cells(4, 2).Value = ""
            
        wb.Save()
        wb.Close()
        print("Successfully removed warning text!")
    except Exception as e:
        print("Error:", e)
    finally:
        excel.Quit()
