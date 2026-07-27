import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.040_20260702_1507.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "wsMain.Range(\"A2:K20\").ClearContents" in line:
        new_lines.append("wsMain.Range(\"A1:Z1\").UnMerge\n")
        new_lines.append("wsMain.Range(\"A1:Z1\").ClearContents\n")
        new_lines.append(line)
    elif "wsMain.Range(\"F1:M1\").Merge" in line:
        new_lines.append(line.replace("F1:M1", "G1:P1"))
    elif "wsMain.Range(\"F1\").Value = GetActiveAgencyName()" in line:
        new_lines.append(line.replace("F1", "G1"))
    elif "wsMain.Range(\"F1\").Font.Size = 28" in line:
        new_lines.append(line.replace("F1", "G1"))
    elif "wsMain.Range(\"F1\").Font.Bold = True" in line:
        new_lines.append(line.replace("F1", "G1"))
    elif "wsMain.Range(\"F1\").Font.Color =" in line:
        new_lines.append(line.replace("F1", "G1"))
    elif "wsMain.Range(\"F1\").HorizontalAlignment =" in line:
        new_lines.append(line.replace("F1", "G1"))
    elif "wsMain.Range(\"F1\").VerticalAlignment =" in line:
        new_lines.append(line.replace("F1", "G1"))
    elif "Attribute VB_Name = \"Goren_Claude_V2_040\"" in line:
        new_lines.append(line.replace("V2_040", "V2_041"))
    elif "Private Const APP_VERSION As String = \"2.040\"" in line:
        new_lines.append(line.replace("2.040", "2.041"))
    elif "VERSION: V2.040" in line:
        new_lines.append(line.replace("2.040", "2.041"))
    elif "Private Const FORCE_DEMO_MODE As Boolean = False" in line:
        new_lines.append(line.replace("False", "True"))
    else:
        new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.041.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.041 with correct edits")
