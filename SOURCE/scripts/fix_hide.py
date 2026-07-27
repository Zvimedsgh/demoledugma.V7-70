import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.063.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.064.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_064"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.064\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.064"\n')
    elif 'For Each ws In ThisWorkbook.Worksheets' in line and "HideWorkSheets" in "".join(lines[max(0, i-50):i]):
        new_lines.append(line)
        new_lines.append('        Select Case ws.Name\n')
        new_lines.append('            Case CONTROL_SHEET_NAME(), MANAGEMENT_SHEET_NAME(), SHEET_COMPANIES(), _\n')
        new_lines.append('                 SHEET_AGENTS(), SHEET_BRANCH(), SHEET_MAINBRANCH(), _\n')
        new_lines.append('                 SHEET_TELLERS(), SHEET_MONTHS(), SHEET_SUMMARY()\n')
        new_lines.append('                \n')
        new_lines.append('                ws.Visible = xlSheetVisible\n')
        new_lines.append('            Case Else\n')
        new_lines.append('                ws.Visible = xlSheetVeryHidden\n')
        new_lines.append('        End Select\n')
        skip = True
    elif skip and 'Next ws' in line:
        new_lines.append(line)
        skip = False
    elif not skip:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.064")

