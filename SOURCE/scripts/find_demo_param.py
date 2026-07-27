import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.048.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "DEMO_MODE" in line:
        print(f"[{i}] {lines[i].strip()}")
