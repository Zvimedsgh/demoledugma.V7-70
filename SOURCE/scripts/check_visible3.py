import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.101.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(max(0, 7180), min(len(lines), 7200)):
    print(f"[{i+1}] {lines[i].strip()}")

