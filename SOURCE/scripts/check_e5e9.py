import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.132.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "E5:E9" in line:
        print(f"[{i+1}] {lines[i].strip()}")

