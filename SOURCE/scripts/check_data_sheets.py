import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.155.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "\"DATA_\"" in line or "DATA_" in line:
        if i - lines.index(line) > 500: break
        print(f"[{i+1}] {line.strip()}")

