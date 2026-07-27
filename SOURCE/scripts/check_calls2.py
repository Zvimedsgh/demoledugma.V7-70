import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.196.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "InitRawColumns" in line:
        for j in range(i-2, i+2):
            print(f"[{j+1}] {lines[j].strip()}")
        print("---")

