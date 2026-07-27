import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.115.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(6166, 6185):
    print(f"[{i+1}] {lines[i].strip()}")

