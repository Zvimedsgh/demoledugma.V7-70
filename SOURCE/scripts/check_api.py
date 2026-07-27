import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.185.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for j in range(6120, 6140):
    print(f"[{j+1}] {lines[j].strip()}")

