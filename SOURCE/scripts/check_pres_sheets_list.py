import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub BuildPresentation()" in line:
        in_func = True
    if in_func:
        if i - lines.index("Public Sub BuildPresentation()\n") > 150: break
        if "sheetList" in line or "DeleteSheet" in line or "ExportTotalChart" in line or "ExportCompCharts" in line:
            print(f"[{i+1}] {lines[i].strip()}")
        if "End Sub" in line:
            in_func = False
            break

