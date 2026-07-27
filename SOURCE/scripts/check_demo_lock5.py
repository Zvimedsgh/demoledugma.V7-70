import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.126.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for j in range(7390, 7410):
    print(f"[{j+1}] {lines[j].strip()}")

