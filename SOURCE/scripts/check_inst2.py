import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.194.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "הוראות_תפעול" in line:
        for j in range(i-2, i+4):
            print(f"[{j+1}] {lines[j].strip()}")

