import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.211.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(4800, 4600, -1):
    if "Sub " in lines[i]:
        print(f"Found Sub: {lines[i].strip()} at line {i+1}")
        break
