import datetime
import re

def create_v968():
    with open(r'C:\ledugma\DEMO\V9.66_CopyPaste.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.68 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.66", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.66", f"'          VERSION 9.68")
    code = code.replace("' CHANGES IN 9.66:", f"' CHANGES IN 9.68:\n' - Removed wsMain.Protect from UpdatePeriodDropdown to stop locking G7!\n' CHANGES IN 9.66:")
    code = code.replace("V9.66", "V9.68")
    code = code.replace('Public Const SYSTEM_VERSION = "9.66"', 'Public Const SYSTEM_VERSION = "9.68"')

    # Remove wsMain.Protect from CLEAN_EXIT of both UpdatePeriodDropdown and UpdateFilterValueDropdown
    code = code.replace("wsMain.Protect SHEET_PROTECT_PWD, UserInterfaceOnly:=True", "")

    # Ensure EnableEvents is True at the end of A00_SetupMainSheet just in case
    # (It's already there from 9.66, but we'll double check)

    with open(r'C:\ledugma\DEMO\V9.68_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v968()
