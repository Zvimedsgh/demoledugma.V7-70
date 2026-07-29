import datetime
import re

def create_v965():
    with open(r'C:\ledugma\DEMO\V9.64_CopyPaste.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.65 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.64", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.64", f"'          VERSION 9.65")
    code = code.replace("' CHANGES IN 9.64:", f"' CHANGES IN 9.65:\n' - Restored F5 to lblShnShotef instead of sug machpil\n' CHANGES IN 9.64:")
    code = code.replace("V9.64", "V9.65")
    code = code.replace('Public Const SYSTEM_VERSION = "9.64"', 'Public Const SYSTEM_VERSION = "9.65"')

    # Fix F5
    old_f5 = 'wsMain.Range("F5").Value = ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1502) & ChrW(1499) & ChrW(1508) & ChrW(1497) & ChrW(1500)'
    new_f5 = 'wsMain.Range("F5").Value = lblShnShotef'
    code = code.replace(old_f5, new_f5)

    with open(r'C:\ledugma\DEMO\V9.65_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v965()
