import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.060.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(6940, 6950):
    print(f"[{i+1}] {lines[i].strip()}")

for i in range(7015, 7025):
    print(f"[{i+1}] {lines[i].strip()}")

