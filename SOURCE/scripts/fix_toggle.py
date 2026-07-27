import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.065.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.066.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_func = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_066"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.066\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.066"\n')
    elif "Public Sub ToggleHiddenSheets()" in line:
        in_func = True
        new_lines.append(line)
        new_lines.append("    Dim ws As Worksheet\n")
        new_lines.append("    Dim anyOtherVisible As Boolean\n")
        new_lines.append("    anyOtherVisible = False\n")
        new_lines.append("    \n")
        new_lines.append("    For Each ws In ThisWorkbook.Worksheets\n")
        new_lines.append("        Select Case ws.Name\n")
        new_lines.append("            Case CONTROL_SHEET_NAME(), MANAGEMENT_SHEET_NAME(), SHEET_COMPANIES(), _\n")
        new_lines.append("                 SHEET_AGENTS(), SHEET_BRANCH(), SHEET_MAINBRANCH(), _\n")
        new_lines.append("                 SHEET_TELLERS(), SHEET_MONTHS(), SHEET_SUMMARY()\n")
        new_lines.append("                 ' These are always visible, so ignore them\n")
        new_lines.append("            Case Else\n")
        new_lines.append("                 If ws.Visible = xlSheetVisible Then\n")
        new_lines.append("                     anyOtherVisible = True\n")
        new_lines.append("                     Exit For\n")
        new_lines.append("                 End If\n")
        new_lines.append("        End Select\n")
        new_lines.append("    Next ws\n")
        new_lines.append("    \n")
        new_lines.append("    If anyOtherVisible Then\n")
        new_lines.append("        HideWorkSheets\n")
        new_lines.append("    Else\n")
        new_lines.append("        ShowHiddenSheets\n")
        new_lines.append("    End If\n")
        new_lines.append("End Sub\n")
    elif in_func and "End Sub" in line:
        in_func = False
    elif not in_func:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.066")

