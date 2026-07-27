import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.155.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Private Sub ExportCompCharts(" in line or "Private Sub ExportTotalChart(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 500: break
        if "Delete" in line or "Alerts" in line:
            print(f"[{i+1}] {line.strip()}")
        if "End Sub" in line:
            in_func = False

