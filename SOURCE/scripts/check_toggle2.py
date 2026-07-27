import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.185.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

found = False
for i, line in enumerate(lines):
    if "ignoreSheet = False" in line:
        found = True
        for j in range(i-5, i+15):
            print(f"[{j+1}] {lines[j].strip()}")
        break
        
if not found:
    print("Could not find ignoreSheet")
