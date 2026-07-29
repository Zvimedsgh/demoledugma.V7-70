import re
import datetime

def create_v972():
    with open(r'C:\ledugma\DEMO\V9.71_Final.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    # Update version
    code = code.replace("9.71", "9.72")
    code = code.replace("V9.71", "V9.72")

    # Change named ranges to be highly unique
    code = code.replace("lst_half_year", "sys_lst_half_year")
    code = code.replace("lst_quarter", "sys_lst_quarter")
    code = code.replace("lst_month", "sys_lst_month")

    # Fix EnsureNamedRanges to unprotect the management sheet
    ensure_old = """Private Sub EnsureNamedRanges()
    Dim wsMgmt As Worksheet
    Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())"""

    ensure_new = """Private Sub EnsureNamedRanges()
    Dim wsMgmt As Worksheet
    Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    
    On Error Resume Next
    wsMgmt.Unprotect SHEET_PROTECT_PWD
    wsMgmt.Unprotect ""
    On Error GoTo 0"""

    code = code.replace(ensure_old, ensure_new)

    # Protect at end of EnsureNamedRanges
    end_ensure_old = """    wsMgmt.Cells(10, 18).Value = "10"
    wsMgmt.Cells(11, 18).Value = "11"
    wsMgmt.Cells(12, 18).Value = "12"
End Sub"""

    end_ensure_new = """    wsMgmt.Cells(10, 18).Value = "10"
    wsMgmt.Cells(11, 18).Value = "11"
    wsMgmt.Cells(12, 18).Value = "12"
    
    wsMgmt.Protect SHEET_PROTECT_PWD, UserInterfaceOnly:=True
End Sub"""

    code = code.replace(end_ensure_old, end_ensure_new)

    with open(r'C:\ledugma\DEMO\V9.72_Final.txt', 'w', encoding='windows-1255') as f:
        f.write(code)
        
    print("V9.72_Final.txt created successfully.")

if __name__ == '__main__':
    create_v972()
