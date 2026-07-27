import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.069.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.070.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_show_hidden = False
skip = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_070"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.070\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.070"\n')
    elif "Public Sub ShowHiddenSheets()" in line:
        in_show_hidden = True
        new_lines.append(line)
    elif in_show_hidden and "If ws.Name <> MATACH_SHEET_NAME() Then" in line:
        new_lines.append(line)
        new_lines.append('            If IsParameterSheet(ws.Name) Then\n')
        new_lines.append('                If ws.Visible <> xlSheetHidden Then ws.Visible = xlSheetHidden\n')
        new_lines.append('            Else\n')
        new_lines.append('                ws.Visible = xlSheetVisible\n')
        new_lines.append('            End If\n')
        skip = True
    elif in_show_hidden and skip and "End If" in line:
        skip = False
    elif in_show_hidden and "End Sub" in line:
        in_show_hidden = False
        new_lines.append(line)
    elif not skip:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.070")

