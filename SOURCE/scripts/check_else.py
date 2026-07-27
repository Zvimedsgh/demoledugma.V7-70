import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.145.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsMain.Range(\"A1\").Font.Color = RGB(200, 0, 0)" in line:
        for j in range(i, min(len(lines), i+6)):
            print(f"[{j+1}] {lines[j].strip()}")
        break

