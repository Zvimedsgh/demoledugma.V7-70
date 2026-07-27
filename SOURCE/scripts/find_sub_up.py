import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(5580, 5600):
    if "Sub " in lines[i]:
        print(f"[{i}] {lines[i].strip()}")
