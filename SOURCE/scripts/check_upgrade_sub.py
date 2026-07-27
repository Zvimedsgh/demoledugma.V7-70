import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.203.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "4590 wsParams.Cells" in line:
        for j in range(i-50, i):
            if " Sub " in lines[j] or " Function " in lines[j]:
                print(f"[{j+1}] {lines[j].strip()}")
        break
