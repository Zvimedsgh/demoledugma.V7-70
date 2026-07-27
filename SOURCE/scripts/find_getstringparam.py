import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.050_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Function GetStringParameter" in line:
        for j in range(i, i+15):
            print(f"[{j}] {lines[j].strip()}")
        break
