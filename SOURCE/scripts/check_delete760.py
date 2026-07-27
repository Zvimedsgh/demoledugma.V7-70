import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.120.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(760, 785):
    print(f"[{i+1}] {lines[i].strip()}")

