import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Private Function H_SET_CLIENTS(" in line or "Public Function H_SET_CLIENTS(" in line:
        for j in range(i, i+3):
            print(f"[{j+1}] {lines[j].strip()}")

