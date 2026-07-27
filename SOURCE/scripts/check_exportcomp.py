import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Private Sub ExportCompCharts(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 100: pass
        if "outSheetName =" in line or "DeleteSheetIfExists" in line:
            print(f"[{i+1}] {lines[i].strip()}")
        if "End Sub" in line:
            in_func = False

