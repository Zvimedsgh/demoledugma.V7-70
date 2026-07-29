import datetime

def create_v960():
    with open(r'C:\ledugma\DEMO\V9.59_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.60 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.59", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.59", f"'          VERSION 9.60")
    code = code.replace("' CHANGES IN 9.59:", f"' CHANGES IN 9.60:\n' - Fixed A00_SetupMainSheet crash due to invalid FormatConditions formula\n' CHANGES IN 9.59:")
    code = code.replace("V9.59", "V9.60")
    code = code.replace('Public Const SYSTEM_VERSION = "9.59"', 'Public Const SYSTEM_VERSION = "9.60"')

    # Fix Format Conditions
    # For G7:
    bad_g7 = 'wsMain.Range("G7").FormatConditions.Add Type:=xlExpression, Formula1:="=($G$7=)*($G$6<>)"'
    good_g7 = 'wsMain.Range("G7").FormatConditions.Add Type:=xlExpression, Formula1:="=AND($G$6<>"""", $G$7=""" & selectText & """)"'
    code = code.replace(bad_g7, good_g7)

    # For G10:
    bad_g10 = 'wsMain.Range("G10").FormatConditions.Add Type:=xlExpression, Formula1:="=($G$10=)*($G$9<>)"'
    good_g10 = 'wsMain.Range("G10").FormatConditions.Add Type:=xlExpression, Formula1:="=AND($G$9<>"""", $G$10=""" & selectText & """)"'
    code = code.replace(bad_g10, good_g10)

    # Make sure we add On Error Resume Next around FormatConditions just to be super safe against crashes
    code = code.replace('wsMain.Range("G7").FormatConditions.Delete', 'On Error Resume Next\n    wsMain.Range("G7").FormatConditions.Delete\n    On Error GoTo 0')
    code = code.replace('wsMain.Range("G10").FormatConditions.Delete', 'On Error Resume Next\n    wsMain.Range("G10").FormatConditions.Delete\n    On Error GoTo 0')

    with open(r'C:\ledugma\DEMO\V9.60_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v960()
