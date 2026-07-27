import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.043_20260702_1545.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub UpdateFilterValueDropdown()" in line:
        for j in range(i+30, i+70):
            print(f"[{j}] {lines[j].strip()}")
        break
