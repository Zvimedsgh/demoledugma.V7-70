import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.195.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(709, 715):
    print(f"[{i+1}] {repr(lines[i])}")

