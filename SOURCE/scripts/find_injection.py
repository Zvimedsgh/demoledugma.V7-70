import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.046_20260702_1615.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "VBComponents" in line or "CodeModule" in line:
        print(f"[{i}] {lines[i].strip()}")
