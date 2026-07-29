import winreg
import win32com.client

def enable_vba_trust():
    try:
        # Excel 16.0 is for Office 2016/2019/365
        key_path = r"Software\Microsoft\Office\16.0\Excel\Security"
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        print("Enabled VBA trust in registry.")
    except Exception as e:
        print(f"Failed to set registry: {e}")

def test_vba():
    enable_vba_trust()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Add()
        bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
        
        try:
            wb.VBProject.VBComponents.Import(bas_file)
            print("Import successful.")
            # If we imported successfully, let's try to access the imported module to trigger a compile
            # There is no direct "Compile" method in VBA object model for external automation easily
            # But we can try to call a dummy function or just see if the VBProject is valid
            print("Module components: ", wb.VBProject.VBComponents.Count)
        except Exception as e:
            print(f"Error during import/compile: {e}")
        finally:
            wb.Close(False)
            excel.Quit()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_vba()
