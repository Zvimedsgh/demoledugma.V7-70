import datetime
import re

def create_v962():
    with open(r'C:\ledugma\DEMO\V9.61_CopyPaste.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.62 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.61", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.61", f"'          VERSION 9.62")
    code = code.replace("' CHANGES IN 9.61:", f"' CHANGES IN 9.62:\n' - Fixed A00_SetupMainSheet dropdown crash (replaced rngPeriodType with G6, rngDateType with G8)\n' - Added Application.EnableEvents = True to A00_SetupMainSheet ERR_HANDLER\n' - Fixed selectText missing in FormatConditions in SetupMainSheet\n' CHANGES IN 9.61:")
    code = code.replace("V9.61", "V9.62")
    code = code.replace('Public Const SYSTEM_VERSION = "9.61"', 'Public Const SYSTEM_VERSION = "9.62"')

    # Fix Validation references in A00_SetupMainSheet
    code = code.replace('wsMain.Range("rngPeriodType").Validation', 'wsMain.Range("G6").Validation')
    code = code.replace('wsMain.Range("rngDateType").Validation', 'wsMain.Range("G8").Validation')
    code = code.replace('wsMain.Range("rngPeriodType").Value', 'wsMain.Range("G6").Value')

    # Fix FormatConditions in A00_SetupMainSheet (replace selectText with the actual Hebrew string)
    # The actual string is: ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497) (בחר/י)
    # In Formula1 it should be literally inside quotes.
    # To be safe, we just use the ChrW code string construction.
    old_g7 = 'wsMain.Range("G7").FormatConditions.Add Type:=xlExpression, Formula1:="=AND($G$6<>"""", $G$7=""" & selectText & """)"'
    new_g7 = 'wsMain.Range("G7").FormatConditions.Add Type:=xlExpression, Formula1:="=AND($G$6<>"""", $G$7=""" & ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497) & """)"'
    code = code.replace(old_g7, new_g7)

    old_g10 = 'wsMain.Range("G10").FormatConditions.Add Type:=xlExpression, Formula1:="=AND($G$9<>"""", $G$10=""" & selectText & """)"'
    new_g10 = 'wsMain.Range("G10").FormatConditions.Add Type:=xlExpression, Formula1:="=AND($G$9<>"""", $G$10=""" & ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497) & """)"'
    code = code.replace(old_g10, new_g10)

    # Add EnableEvents = True to ERR_HANDLER of A00_SetupMainSheet
    # Find A00_SetupMainSheet ERR_HANDLER
    err_block = """ERR_HANDLER:
    MsgBoxU wsMgmt.Cells(10, 19).Value & Err.Description, vbCritical"""
    
    new_err_block = """ERR_HANDLER:
    Application.EnableEvents = True
    MsgBoxU wsMgmt.Cells(10, 19).Value & Err.Description, vbCritical"""
    
    # We will use regex to find and replace the specific block inside A00_SetupMainSheet
    code = re.sub(r'ERR_HANDLER:\s*\d*\s*MsgBoxU wsMgmt\.Cells\(10, 19\)\.Value & Err\.Description, vbCritical',
                  r'ERR_HANDLER:\n    Application.EnableEvents = True\n    MsgBoxU wsMgmt.Cells(10, 19).Value & Err.Description, vbCritical',
                  code, count=1)

    # Add it to UpdatePeriodDropdown and UpdateFilterValueDropdown ERR_HANDLER just in case
    code = code.replace('MsgBoxU wsMgmt.Cells(10, 19).Value & Err.Description, vbCritical',
                        'Application.EnableEvents = True\n    MsgBoxU wsMgmt.Cells(10, 19).Value & Err.Description, vbCritical')

    with open(r'C:\ledugma\DEMO\V9.62_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v962()
