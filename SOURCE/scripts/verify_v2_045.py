import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "INDIRECT" in line or "map_period" in line or "Exit Sub" in line:
        if "UpdatePeriodDropdown" in lines[i-1] or "UpdateFilterValueDropdown" in lines[i-1] or "INDIRECT" in line or "map_period" in line:
            print(f"[{i}] {lines[i].strip()}")
