import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.211.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if re.match(r'^\s*Stop\s*$', line) or re.match(r'^\s*\d+\s+Stop\s*$', line):
        print(f"[{i+1}] {line.strip()}")

