import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.103.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Sub NavSettings_Home()" in line:
        in_func = True
        print(f"Line {i+1}: {line.strip()}")
    elif in_func and "End Sub" in line:
        print(f"Line {i+1}: {line.strip()}")
        break
    elif in_func:
        print(f"Line {i+1}: {line.strip()}")

