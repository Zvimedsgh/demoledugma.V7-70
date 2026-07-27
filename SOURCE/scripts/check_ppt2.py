import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.211.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Set ppApp = GetObject" in line:
        for j in range(i-20, i+200):
            try:
                print(f"[{j+1}] {lines[j].strip()}")
            except:
                pass
        break
