import datetime
import re

def create_v964():
    with open(r'C:\ledugma\DEMO\V9.63_CopyPaste.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.64 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.63", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.63", f"'          VERSION 9.64")
    code = code.replace("' CHANGES IN 9.63:", f"' CHANGES IN 9.64:\n' - Added explicit wsMain.Unprotect/Protect in UpdatePeriodDropdown to prevent Validation.Add crash if sheet is protected\n' - Unlocked G6:G12 during SetupMainSheet\n' CHANGES IN 9.63:")
    code = code.replace("V9.63", "V9.64")
    code = code.replace('Public Const SYSTEM_VERSION = "9.63"', 'Public Const SYSTEM_VERSION = "9.64"')

    # Fix UpdatePeriodDropdown to Unprotect and Protect
    # Find the start of UpdatePeriodDropdown
    upd_start = """30      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())

40      periodType = Trim$(CStr(wsMain.Range("G6").Value2))"""
    upd_start_new = """30      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        wsMain.Unprotect SHEET_PROTECT_PWD

40      periodType = Trim$(CStr(wsMain.Range("G6").Value2))"""
    code = code.replace(upd_start, upd_start_new)
    
    # Same for UpdateFilterValueDropdown
    ufv_start = """30      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())

40      filterType = Trim$(CStr(wsMain.Range("G9").Value2))"""
    ufv_start_new = """30      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        wsMain.Unprotect SHEET_PROTECT_PWD

40      filterType = Trim$(CStr(wsMain.Range("G9").Value2))"""
    code = code.replace(ufv_start, ufv_start_new)

    # Protect at CLEAN_EXIT of UpdatePeriodDropdown
    code = code.replace("CLEAN_EXIT:\n\n210     Exit Sub", "CLEAN_EXIT:\n        wsMain.Protect SHEET_PROTECT_PWD, UserInterfaceOnly:=True\n210     Exit Sub")

    # Add Locked = False to A00_SetupMainSheet
    code = code.replace("wsMain.Unprotect Password:=SHEET_PASSWORD()", "wsMain.Unprotect Password:=SHEET_PASSWORD()\n    wsMain.Range(\"G6:G12\").Locked = False")

    with open(r'C:\ledugma\DEMO\V9.64_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v964()
