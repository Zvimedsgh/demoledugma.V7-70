import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045_20260702_1605.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Dim " in line and 3280 <= i <= 3600:
        print(f"[{i}] {lines[i].strip()}")
