import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.103.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "A1" in line and ("Goto" in line or "Select" in line or "Activate" in line):
        print(f"Line {i+1}: {line.strip()}")
        # print the subroutine context
        for j in range(i, -1, -1):
            if "Sub " in lines[j] or "Function " in lines[j]:
                print(f"   Context: {lines[j].strip()}")
                break
        print("-" * 40)

