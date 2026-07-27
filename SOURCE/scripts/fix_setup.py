import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.070.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.071.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_071"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.071\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.071"\n')
    elif "wsMain.Range(\"A2:K20\").ClearContents" in line:
        new_lines.append(line)
        new_lines.append("    ' Clear all existing rounded rectangle buttons to avoid duplicates if run multiple times\n")
        new_lines.append("    Dim shpBtn As Shape\n")
        new_lines.append("    For Each shpBtn In wsMain.Shapes\n")
        new_lines.append("        If shpBtn.Type = 5 Then ' msoShapeRoundedRectangle\n")
        new_lines.append("            shpBtn.Delete\n")
        new_lines.append("        End If\n")
        new_lines.append("    Next shpBtn\n")
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.071")

