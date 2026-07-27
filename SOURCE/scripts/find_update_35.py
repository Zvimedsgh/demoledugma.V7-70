import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub UpdatePeriodDropdown" in line or "UpdateFilterDropdown" in line:
        print(f"[{i}] {line.strip()}")
