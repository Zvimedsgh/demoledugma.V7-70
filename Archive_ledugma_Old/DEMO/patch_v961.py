import datetime

def create_v961():
    with open(r'C:\ledugma\DEMO\V9.60_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.61 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.60", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.60", f"'          VERSION 9.61")
    code = code.replace("' CHANGES IN 9.60:", f"' CHANGES IN 9.61:\n' - Prevent overwriting J4/K5 currency labels in SetupMainSheet if they exist\n' CHANGES IN 9.60:")
    code = code.replace("V9.60", "V9.61")
    code = code.replace('Public Const SYSTEM_VERSION = "9.60"', 'Public Const SYSTEM_VERSION = "9.61"')

    # Fix J4/J5 population to only happen if empty
    code = code.replace(
        'wsMain.Range("J4").Value = ChrW(1491) & ChrW(1493) & ChrW(1500) & ChrW(1512)',
        'If wsMain.Range("J4").Value = "" Then wsMain.Range("J4").Value = ChrW(1491) & ChrW(1493) & ChrW(1500) & ChrW(1512)'
    )
    code = code.replace(
        'wsMain.Range("J5").Value = ChrW(1488) & ChrW(1497) & ChrW(1512) & ChrW(1493)',
        'If wsMain.Range("J5").Value = "" Then wsMain.Range("J5").Value = ChrW(1488) & ChrW(1497) & ChrW(1512) & ChrW(1493)'
    )
    code = code.replace(
        'wsMain.Range("K4").Value = 3.996',
        'If wsMain.Range("K4").Value = "" Then wsMain.Range("K4").Value = 3.996'
    )
    code = code.replace(
        'wsMain.Range("K5").Value = 3.4114',
        'If wsMain.Range("K5").Value = "" Then wsMain.Range("K5").Value = 3.4114'
    )

    with open(r'C:\ledugma\DEMO\V9.61_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v961()
