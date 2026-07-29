import datetime

def create_v959():
    with open(r'C:\ledugma\DEMO\V9.58_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.59 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.58", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.58", f"'          VERSION 9.59")
    code = code.replace("' CHANGES IN 9.58:", f"' CHANGES IN 9.59:\n' - Fixed ClearHomePageSelection button assignment\n' - Added safety EnableEvents reset\n' CHANGES IN 9.58:")
    code = code.replace("V9.58", "V9.59")
    code = code.replace('Public Const SYSTEM_VERSION = "9.58"', 'Public Const SYSTEM_VERSION = "9.59"')

    # Fix button macro assignment
    code = code.replace('btnClearSel.OnAction = "ClearSelection"', 'btnClearSel.OnAction = "ClearHomePageSelection"')

    with open(r'C:\ledugma\DEMO\V9.59_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v959()
