import re
import io

def shift_match(m):
    col1 = m.group(1)
    row1 = int(m.group(2)) + 1
    if m.group(3):
        col2 = m.group(3)
        row2 = int(m.group(4)) + 1
        return f'{col1}{row1}:{col2}{row2}'
    else:
        return f'{col1}{row1}'

def shift_range_string(s):
    return re.sub(r'([A-Z])([0-9]+)(?::([A-Z])([0-9]+))?', shift_match, s)

with open(r'C:\ledugma\DEMO\modDemoReports_V9.36.bas', encoding='windows-1255', errors='ignore') as f:
    lines = f.readlines()

new_lines = []
for line in lines:
    orig_line = line
    # Shift wsMain.Range(...)
    if 'wsMain.Range("' in line:
        line = re.sub(r'(wsMain\.Range\(")([^"]+)("\))', lambda m: m.group(1) + shift_range_string(m.group(2)) + m.group(3), line)
    
    # Shift ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range(...)
    if 'ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("' in line:
        line = re.sub(r'(ThisWorkbook\.Worksheets\(CONTROL_SHEET_NAME\(\)\)\.Range\(")([^"]+)("\))', lambda m: m.group(1) + shift_range_string(m.group(2)) + m.group(3), line)
    
    # Shift RefersTo:="='" & wsMain.Name & "'!...
    if 'RefersTo:="=\'" & wsMain.Name & "\'!$' in line:
        line = re.sub(r'(RefersTo:="=\'" & wsMain\.Name & "\'!\$)([A-Z])\$([0-9]+)', lambda m: f"{m.group(1)}{m.group(2)}${int(m.group(3))+1}", line)
        
    # Shift wsMain.Rows(...)
    if 'wsMain.Rows("' in line:
        line = re.sub(r'(wsMain\.Rows\(")([0-9]+):([0-9]+)("\))', lambda m: f"{m.group(1)}{int(m.group(2))+1}:{int(m.group(3))+1}{m.group(4)}", line)
        
    new_lines.append(line)

with io.open(r'C:\ledugma\DEMO\modDemoReports_V9.37.bas', 'w', encoding='windows-1255') as f:
    for line in new_lines:
        f.write(line)
