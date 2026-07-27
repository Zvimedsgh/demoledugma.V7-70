import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.170.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "btnNewClients" in line:
        for j in range(max(0, i-5), min(len(lines), i+15)):
            print(f"[{j+1}] {lines[j].strip()}")

