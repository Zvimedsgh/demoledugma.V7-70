import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.199.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if " Sub " in line and ("Filter" in line or "סנן" in line):
        print(f"[{i+1}] {line.strip()}")

