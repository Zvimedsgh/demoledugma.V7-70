import datetime
import re

def create_v963():
    with open(r'C:\ledugma\DEMO\V9.62_CopyPaste.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.63 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.62", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.62", f"'          VERSION 9.63")
    code = code.replace("' CHANGES IN 9.62:", f"' CHANGES IN 9.63:\n' - Fixed missing dropdowns: converted all remaining rng references to G6/G7/G10 in Setup and Handlers\n' CHANGES IN 9.62:")
    code = code.replace("V9.62", "V9.63")
    code = code.replace('Public Const SYSTEM_VERSION = "9.62"', 'Public Const SYSTEM_VERSION = "9.63"')

    # Fix UpdatePeriodDropdown references
    code = code.replace('wsMain.Range("rngPeriodValue")', 'wsMain.Range("G7")')
    
    # Fix UpdateFilterValueDropdown references
    code = code.replace('wsMain.Range("rngFilterValue")', 'wsMain.Range("G10")')

    # Fix any remaining rngPeriodType/rngDateType in SetupMainSheet
    code = code.replace('wsMain.Range("rngPeriodType")', 'wsMain.Range("G6")')
    code = code.replace('wsMain.Range("rngDateType")', 'wsMain.Range("G8")')

    # Also make absolutely sure the G dropdowns are UNLOCKED for the user
    # Insert wsMain.Range("G6:G12").Locked = False right after the protection clear
    code = code.replace('wsMain.Cells.Validation.Delete', 'wsMain.Cells.Validation.Delete\n    wsMain.Range("G6:G12").Locked = False')

    with open(r'C:\ledugma\DEMO\V9.63_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v963()
