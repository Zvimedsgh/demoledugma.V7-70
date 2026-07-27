import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Private Sub ExportCompCharts" in line or "Private Sub ExportTotalChart" in line or "Public Sub ApplyCorrectionsAndBuildReports" in line or "Public Sub BuildPresentation" in line:
        in_func = line.strip()
    if in_func:
        if i - lines.index(line) > 500: pass
        if "Set wsOut =" in line or "wsOut.Name =" in line or ".Delete" in line and "ws" in line:
            print(f"[{i+1}] {in_func}: {lines[i].strip()}")

