import win32com.client
import sys

def test_vba_file(filepath):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Add()
        
        try:
            wb.VBProject.VBComponents.Import(filepath)
            print(f"Import {filepath} successful.")
        except Exception as e:
            print(f"Error during import: {e}")
        finally:
            wb.Close(False)
            excel.Quit()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_vba_file(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.223.bas')
