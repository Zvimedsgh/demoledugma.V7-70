import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Private Sub ExportCompCharts(" in line_strip or "Private Sub ExportTotalChart(" in line_strip:
        in_func = True
    elif in_func and "End Sub" in line_strip:
        in_func = False
    elif in_func:
        if "Const " in line_strip:
            print(f"[{i+1}] {line_strip}")

