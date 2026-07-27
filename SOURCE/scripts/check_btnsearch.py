import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.141.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "btnSearch" in line or "shp.Name =" in line and "Search" in line:
        for j in range(max(0, i-2), i+3):
            print(f"[{j+1}] {lines[j].strip()}")

