import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.071.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(615, 665):
    print(f"[{i+1}] {lines[i].strip()}")

