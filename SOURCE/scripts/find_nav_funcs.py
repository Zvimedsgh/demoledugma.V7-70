import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(7715, 7730):
    if i < len(lines):
        print(f"[{i}] {lines[i].strip()}")
