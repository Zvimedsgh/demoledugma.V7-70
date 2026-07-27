import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.183.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

max_len = 0
for i, line in enumerate(lines):
    if len(line) > max_len:
        max_len = len(line)
        if max_len > 200:
            print(f"Line {i+1} length: {max_len}")

print(f"Max length overall: {max_len}")
