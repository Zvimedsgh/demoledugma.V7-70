import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045_20260702_1605.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "ThisWorkbook.Names.Add \"lst_month\"" in line:
        for j in range(i-5, i+15):
            print(f"[{j}] {lines[j].strip()}")
        break
