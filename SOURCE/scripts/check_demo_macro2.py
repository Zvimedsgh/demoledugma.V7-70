import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.160.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3650, 3700):
    if "Public Sub " in lines[i] or "Private Sub " in lines[i]:
        print(f"[{i+1}] {lines[i].strip()}")

