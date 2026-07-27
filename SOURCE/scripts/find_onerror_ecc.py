import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.052_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Private Sub ExportCompCharts" in line:
        in_func = True
    if in_func and "On Error" in line:
        print(f"[{i}] {lines[i].strip()}")
    if in_func and "End Sub" in line:
        in_func = False
        break
