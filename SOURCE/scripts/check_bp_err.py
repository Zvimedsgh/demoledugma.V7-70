import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.211.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "ERR_HANDLER:" in line and i > 4600 and i < 5100:
        for j in range(i, i+50):
            try:
                print(f"[{j+1}] {lines[j].strip()}")
            except:
                pass
        break
