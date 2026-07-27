import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "DeleteSheetIfExists sheetName" in line:
        for j in range(max(0, i-20), min(len(lines), i+10)):
            print(f"[{j+1}] {lines[j].strip()}")
        print("---")

