import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "dictCustComp(" in line:
        for j in range(max(0, i-5), min(len(lines), i+5)):
            print(f"[{j+1}] {lines[j].strip()}")
        break

