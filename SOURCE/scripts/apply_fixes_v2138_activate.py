import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "    wsSearch.Activate\n" == line:
        if "wsSearch.Visible = xlSheetVisible" not in lines[i-1]:
            lines.insert(i, "    wsSearch.Visible = xlSheetVisible\n")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Activate fixed.")
