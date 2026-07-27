import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.193.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsLists.Cells(1, 2).Value = ChrW(1496)" in line or "wsLists.Cells(1, 2).Value" in line:
        for j in range(i-2, i+20):
            print(f"[{j+1}] {lines[j].strip()}")
        break

