import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.156.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Sub ShowHiddenSheets" in line or "Sub HideWorkSheets" in line:
        for j in range(i, i+30):
            print(f"[{j+1}] {lines[j].strip()}")
        print("-" * 20)

