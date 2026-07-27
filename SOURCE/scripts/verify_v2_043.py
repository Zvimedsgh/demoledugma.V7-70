import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.043.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsMain.Range(\"A1:F1\").Merge" in line or "RowHeight = 120" in line:
        print(f"[{i}] {lines[i].strip()}")
