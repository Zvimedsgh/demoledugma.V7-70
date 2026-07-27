import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.196.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if chr(1513) + chr(1490) + chr(1497) + chr(1488) + chr(1492) in line:  # "שגיאה"
        print(f"[{i+1}] {line.strip()}")

