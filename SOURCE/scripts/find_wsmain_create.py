import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.044_20260702_1555.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsMain.Delete" in line or "wsMain Is Nothing Then" in line or "Worksheets.Add" in line:
        for j in range(i-2, i+5):
            print(f"[{j}] {lines[j].strip()}")
        break
