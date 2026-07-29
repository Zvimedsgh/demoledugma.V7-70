import win32com.client
import os

def test_macro():
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    wb_path = r"C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70_patched.xlsm"
    
    try:
        wb = xl.Workbooks.Open(wb_path)
        
        # Read V9.63
        with open(r'C:\ledugma\DEMO\V9.63_CopyPaste.txt', 'r', encoding='windows-1255') as f:
            code = f.read()
            
        # Inject into modMain
        for comp in wb.VBProject.VBComponents:
            if comp.Name == "modMain" or comp.Name == "modReports":
                # Assuming the module name is modReports or modMain
                # Let's just delete the lines and insert
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(code)
                print(f"Injected into {comp.Name}")
                break
                
        # Run A00_SetupMainSheet
        print("Running A00_SetupMainSheet...")
        xl.Application.Run("A00_SetupMainSheet")
        print("A00_SetupMainSheet ran successfully!")
        
        # Test changing G6
        ws = wb.Sheets("ראשי")
        print("Current G6:", ws.Range("G6").Value)
        print("Current G7:", ws.Range("G7").Value)
        
        # Trigger Worksheet_Change by changing G6
        # To do this safely, we can just call UpdatePeriodDropdown directly
        print("Calling UpdatePeriodDropdown directly...")
        ws.Range("G6").Value = "שנתי"
        xl.Application.Run("UpdatePeriodDropdown")
        print("After changing G6 to yearly, G7:", ws.Range("G7").Value)
        
        ws.Range("G6").Value = "חצי שנתי"
        xl.Application.Run("UpdatePeriodDropdown")
        print("After changing G6 to half-yearly, G7:", ws.Range("G7").Value)
        try:
            print("G7 Validation Formula:", ws.Range("G7").Validation.Formula1)
        except Exception as e:
            print("G7 Validation error:", e)

    except Exception as e:
        print(f"Error: {e}")
    finally:
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        xl.Quit()

if __name__ == "__main__":
    test_macro()
