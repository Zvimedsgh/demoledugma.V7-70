import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.038_20260702_1440.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "IndentLevel" in line or "Spaces" in line or "נקודות" in line:
        print(f"[{i}] {lines[i].strip()}")
