import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Private Function MANAGEMENT_SHEET_NAME" in line:
        print(f"[{i}] {line.strip()}")
        for j in range(i+1, i+3):
            print(f"[{j}] {lines[j].strip()}")
