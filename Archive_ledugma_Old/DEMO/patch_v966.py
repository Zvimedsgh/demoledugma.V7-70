import datetime
import re

def create_v966():
    with open(r'C:\ledugma\DEMO\V9.65_CopyPaste.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.66 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.65", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.65", f"'          VERSION 9.66")
    code = code.replace("' CHANGES IN 9.65:", f"' CHANGES IN 9.66:\n' - Added Application.EnableEvents = True to the successful exit path of A00_SetupMainSheet!\n' CHANGES IN 9.65:")
    code = code.replace("V9.65", "V9.66")
    code = code.replace('Public Const SYSTEM_VERSION = "9.65"', 'Public Const SYSTEM_VERSION = "9.66"')

    # Fix the missing EnableEvents = True before Exit Sub
    code = code.replace("920     Exit Sub", "Application.EnableEvents = True\n920     Exit Sub")

    with open(r'C:\ledugma\DEMO\V9.66_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v966()
