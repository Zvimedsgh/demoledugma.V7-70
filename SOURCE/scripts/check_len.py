import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.182.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if len(line) > 500:
        print(f"Line {i+1} length: {len(line)}")

