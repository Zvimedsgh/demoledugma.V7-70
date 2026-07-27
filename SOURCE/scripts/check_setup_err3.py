import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.072.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub A00_SetupMainSheet()" in line_strip:
        in_func = True
    elif in_func and "End Sub" in line_strip:
        break
    elif in_func and "10 " in line_strip:
        print(f"[{i+1}] {line_strip}")
    elif in_func and "ERR_HANDLER:" in line_strip:
        for j in range(i, i+10):
            print(f"[{j+1}] {lines[j].strip()}")
        break

