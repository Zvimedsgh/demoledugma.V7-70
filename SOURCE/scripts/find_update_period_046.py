import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.046_20260702_1615.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub UpdatePeriodDropdown" in line:
        for j in range(i, i+30):
            print(f"[{j}] {lines[j].strip()}")
        break
