import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.199.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(350, len(lines)):
    if " Sub " in lines[i] or " Function " in lines[i]:
        print(f"First Sub/Function after 350 is at {i+1}: {lines[i].strip()}")
        break

