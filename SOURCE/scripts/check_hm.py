import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.207.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Private Function H_SET_MESSAGES" in line:
        for j in range(i, i+10):
            try:
                print(f"[{j+1}] {lines[j].strip()}")
            except:
                pass
        break
