import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.125.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "shp.Delete" in line or "Shapes(" in line and ".Delete" in line:
        for j in range(max(0, i-2), min(len(lines), i+3)):
            print(f"[{j+1}] {lines[j].strip()}")
        print("-" * 20)

