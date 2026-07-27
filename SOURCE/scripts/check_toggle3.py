import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.184.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Sub ToggleHiddenSheets" in line:
        for j in range(i, i+25):
            print(f"[{j+1}] {lines[j].strip()}")
        break

