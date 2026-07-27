import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Private Sub ExportCompCharts" in line:
        print(f"[{i+1}] {line.strip()}")
        for j in range(i+1, min(len(lines), i+30)):
            if "outSheetName" in line or "wsOut" in lines[j]:
                print(f"[{j+1}] {lines[j].strip()}")

