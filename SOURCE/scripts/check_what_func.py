import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.174.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(7410, 7450):
    print(f"[{i+1}] {lines[i].strip()}")

