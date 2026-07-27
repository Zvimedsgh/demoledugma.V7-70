import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.041.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "wsMain.Range(\"C1:K1\").Merge" in line:
        new_lines.append(line.replace("C1:K1", "G1:P1"))
    elif "wsMain.Range(\"C1\").Value = GetActiveAgencyName()" in line:
        new_lines.append(line.replace("C1", "G1"))
    elif "wsMain.Range(\"C1\").Font.Size = 28" in line:
        new_lines.append(line.replace("C1", "G1"))
    elif "wsMain.Range(\"C1\").Font.Bold = True" in line:
        new_lines.append(line.replace("C1", "G1"))
    elif "wsMain.Range(\"C1\").Font.Color =" in line:
        new_lines.append(line.replace("C1", "G1"))
    elif "wsMain.Range(\"C1\").HorizontalAlignment =" in line:
        new_lines.append(line.replace("C1", "G1"))
    elif "wsMain.Range(\"C1\").VerticalAlignment =" in line:
        new_lines.append(line.replace("C1", "G1"))
    else:
        new_lines.append(line)

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Title shifted correctly in V2.041")
