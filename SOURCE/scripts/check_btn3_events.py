import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.068.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in (4233, 4240, 4549):
    print(f"[{i+1}] {lines[i].strip()}")

