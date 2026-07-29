with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

import re

# We will just fix line 5626
for i, line in enumerate(lines):
    if line.strip().startswith("insuredRef |"):
        lines[i] = "' " + line.lstrip()
        print(f"Fixed line {i+1}")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.writelines(lines)
