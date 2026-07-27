import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.042_20260702_1530.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "wsMain.Range(\"G1:P1\").Merge" in line:
        new_lines.append(line.replace("G1:P1", "A1:F1"))
    elif "wsMain.Range(\"G1\").Value = GetActiveAgencyName()" in line:
        new_lines.append(line.replace("G1", "A1"))
    elif "wsMain.Range(\"G1\").Font.Size = 28" in line:
        new_lines.append(line.replace("G1", "A1"))
    elif "wsMain.Range(\"G1\").Font.Bold = True" in line:
        new_lines.append(line.replace("G1", "A1"))
    elif "wsMain.Range(\"G1\").Font.Color = RGB(200, 0, 0)" in line:
        new_lines.append(line.replace("G1", "A1"))
    elif "wsMain.Range(\"G1\").HorizontalAlignment = xlCenter" in line:
        new_lines.append(line.replace("G1", "A1"))
    elif "wsMain.Range(\"G1\").VerticalAlignment = xlCenter" in line:
        new_lines.append(line.replace("G1", "A1"))
    elif "wsMain.Rows(\"1\").RowHeight = 60" in line:
        new_lines.append(line.replace("60", "120"))
    elif "Attribute VB_Name = \"Goren_Claude_V2_042\"" in line:
        new_lines.append(line.replace("V2_042", "V2_043"))
    elif "Private Const APP_VERSION As String = \"2.042\"" in line:
        new_lines.append(line.replace("2.042", "2.043"))
    elif "VERSION: V2.042" in line:
        new_lines.append(line.replace("2.042", "2.043"))
    else:
        new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.043.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.043")
