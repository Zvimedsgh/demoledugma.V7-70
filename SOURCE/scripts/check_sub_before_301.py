import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.199.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(0, 301):
    if " Sub " in lines[i] or " Function " in lines[i]:
        print(f"Found Sub/Function at {i+1}: {lines[i].strip()}")

