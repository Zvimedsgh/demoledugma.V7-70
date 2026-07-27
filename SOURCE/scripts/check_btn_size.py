import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.123.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3890, 3935):
    if "Set btnSearch =" in lines[i]:
        for j in range(max(0, i-5), i+15):
            print(f"[{j+1}] {lines[j].strip()}")
        break

