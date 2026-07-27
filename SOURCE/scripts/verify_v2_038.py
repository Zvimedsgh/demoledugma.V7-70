import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.038.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3840, 3865):
    print(f"[{i}] {lines[i].strip()}")

for i in range(3880, 3900):
    print(f"[{i}] {lines[i].strip()}")
