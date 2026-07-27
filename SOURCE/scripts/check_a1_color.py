import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.105.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for j in range(3840, 3860):
    print(f"[{j+1}] {lines[j].strip()}")

