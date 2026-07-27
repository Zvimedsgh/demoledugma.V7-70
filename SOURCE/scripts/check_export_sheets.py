import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Private Sub ExportTotalChart" in line or "Private Sub ExportCompCharts" in line:
        in_func = True
    if in_func:
        if "Set wsOut =" in line or ".Delete" in line and "ws" in line:
            print(f"[{i+1}] {lines[i].strip()}")
        if "End Sub" in line:
            in_func = False

