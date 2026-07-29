with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
lines = text.split('\n')
for i, line in enumerate(lines):
    if 'SetAllTabColors' in line and 'Sub SetAllTabColors' not in line:
        print(f"Line {i+1}: {line.strip()}")
        for j in range(i-10, i):
            if 'Sub ' in lines[j]:
                print(f"  Called from: {lines[j].strip()}")
