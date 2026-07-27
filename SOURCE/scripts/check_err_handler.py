import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.206.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "ERR_HANDLER:" in line:
        for j in range(i, i+30):
            try:
                print(f"[{j+1}] {lines[j].strip()}")
            except:
                pass
        break
