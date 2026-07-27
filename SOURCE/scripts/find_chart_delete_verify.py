import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.055.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "tmpWs.ChartObjects.Delete" in line:
        for j in range(i-3, i+4):
            print(f"[{j}] {lines[j].strip()}")
        break
