import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Function CONTROL_SHEET_NAME" in line or "Function MANAGEMENT_SHEET" in line or "Function SHEET_" in line:
        print(line.strip())

