import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045_20260702_1605.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Dim backupG5" in line:
        print(f"Line {i}: {line.strip()}")
