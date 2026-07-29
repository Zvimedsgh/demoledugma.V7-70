import datetime

def create_v958():
    with open(r'C:\ledugma\DEMO\V9.57_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.58 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.49 (2026-06-29)", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.49", f"'          VERSION 9.58")
    code = code.replace("' CHANGES IN 9.49:", f"' CHANGES IN 9.58:\n' - Fixed Event Handlers and Formatting\n' - Fixed Version strings\n' CHANGES IN 9.49:")
    code = code.replace("' V9.49 (2026-06-29)", f"' V9.58 ({date_str} {time_str})")
    code = code.replace('Public Const SYSTEM_VERSION = "9.49"', 'Public Const SYSTEM_VERSION = "9.58"')

    with open(r'C:\ledugma\DEMO\V9.58_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v958()
