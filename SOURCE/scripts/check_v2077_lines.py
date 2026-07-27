import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.077.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "1710 " in line:
        print(f"[{i+1}] {line.strip()}")
    elif "1700 " in line:
        print(f"[{i+1}] {line.strip()}")

