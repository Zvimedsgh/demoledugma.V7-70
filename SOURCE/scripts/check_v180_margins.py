import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.180.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "MarginRight = 10" in line:
        print(f"[{i-1}] {lines[i-2].strip()}")
        print(f"[{i}] {lines[i-1].strip()}")
        print(f"[{i+1}] {line.strip()}")

