import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.114.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(2460, 2480):
    print(f"[{i+1}] {lines[i].strip()}")

for i in range(2650, 2700):
    print(f"[{i+1}] {lines[i].strip()}")

