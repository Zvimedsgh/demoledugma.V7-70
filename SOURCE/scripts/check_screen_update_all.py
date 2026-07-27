import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.114.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "ScreenUpdating" in line and "True" in line:
        # Check if it's commented out
        if not re.match(r'^\s*\'', line):
            print(f"[{i+1}] {line.strip()}")

