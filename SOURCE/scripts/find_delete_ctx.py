import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.029.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(730, 760):
    print(f"[{i}] {lines[i].strip()}")
