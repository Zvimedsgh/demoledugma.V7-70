import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub OpenDataSheet(" in line or "Private Sub OpenDataSheet(" in line or "Sub OpenDataSheet" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 100: pass
        if "Delete" in line and "baseSheetName" in line or "SheetExists" in line:
            for j in range(max(0, i-5), min(len(lines), i+5)):
                print(f"[{j+1}] {lines[j].strip()}")
            break

