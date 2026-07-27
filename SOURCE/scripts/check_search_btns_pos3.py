import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.130.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "btnLeft" in line and "wsSearch.Cells(2, 2).Left" in line:
        for j in range(max(0, i-2), min(len(lines), i+8)):
            print(f"[{j+1}] {lines[j].strip()}")

