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
    install_file = os.path.join(target_dir, "Install_Instructions.xlsx")
    bait_file = os.path.join(target_dir, "Demo_Reports_Syatem_V7.70  (2)3.xlsm")
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(install_file)
        ws = wb.Sheets(1)
        
        # The new link we want to set
        new_url = "https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQBnB0klIujxR4O5fALynO8EAWddhgVppMJI7THxhW3R6fo?e=p3fjlk&download=1"
        
        # Try to find the hyperlink in shapes
        for shp in ws.Shapes:
            try:
                if shp.Hyperlink:
                    shp.Hyperlink.Address = new_url
            except Exception:
                pass
                
        # Also try to find hyperlinks in cells
        for hl in ws.Hyperlinks:
            hl.Address = new_url
            
        # Add the warning text
        ws.Cells(4, 2).Value = "שימו לב: המערכת תרד כעת למחשב שלכם. לאחר ההורדה, פתחו אותה והמשיכו לפי ההוראות."
        ws.Cells(4, 2).Font.Size = 14
        ws.Cells(4, 2).Font.Color = 255 # Red
        ws.Cells(4, 2).Font.Bold = True
            
        # Save as xlsm (FileFormat=52)
        wb.SaveAs(bait_file, FileFormat=52)
        wb.Close()
        print("Successfully updated bait file!")
    except Exception as e:
        print("Error:", e)
    finally:
        excel.Quit()
