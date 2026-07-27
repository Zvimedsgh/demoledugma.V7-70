import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.043_20260702_1545.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip_filter = False
for i, line in enumerate(lines):
    if skip_filter:
        if "390     If valList =" in line:
            skip_filter = False
            new_lines.append(line)
        continue

    # UpdatePeriodDropdown replacements
    if "listName = ChrW(1502) & ChrW(1495)" in line and "ChrW(1492)" in line:
        new_lines.append('100         listName = "=lst_half_year"\n')
    elif "listName = ChrW(1512) & ChrW(1489)" in line and "ChrW(1497)" in line:
        new_lines.append('120         listName = "=lst_quarter"\n')
    elif "listName = ChrW(1497) & ChrW(1504)" in line and "ChrW(1512)" in line:
        new_lines.append('140         listName = "=lst_month"\n')
    
    # UpdateFilterValueDropdown replacements
    elif "320     valList = \"\"" in line:
        new_lines.append('320     ThisWorkbook.Names.Add "lst_temp_filter", wsLists.Range(wsLists.Cells(2, targetCol), wsLists.Cells(lastR, targetCol))\n')
        new_lines.append('330     valList = "=lst_temp_filter"\n')
        skip_filter = True
    
    # Version update
    elif "Attribute VB_Name = \"Goren_Claude_V2_043\"" in line:
        new_lines.append(line.replace("V2_043", "V2_044"))
    elif "Private Const APP_VERSION As String = \"2.043\"" in line:
        new_lines.append(line.replace("2.043", "2.044"))
    elif "VERSION: V2.043" in line:
        new_lines.append(line.replace("2.043", "2.044"))
    else:
        new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.044.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.044 via line-by-line replacement")
