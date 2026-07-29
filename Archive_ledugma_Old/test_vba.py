import win32com.client
import os

def test_compile():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = excel.Workbooks.Add()
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    
    try:
        wb.VBProject.VBComponents.Import(bas_file)
        print("Import successful.")
        
        # We can't explicitly compile via COM easily, but we can try to save or just read the code
        # Actually, if there is a syntax error, it might not error on import, but we can't be sure.
        print("Done importing.")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        wb.Close(False)
        excel.Quit()

if __name__ == "__main__":
    test_compile()
