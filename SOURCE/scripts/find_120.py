import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.042_20260702_1530.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "RowHeight" in line and "120" in line:
        print(f"[{i}] {lines[i].strip()}")
