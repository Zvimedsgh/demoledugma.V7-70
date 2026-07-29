with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
lines = text.split('\n')
for i, line in enumerate(lines):
    if 'ApplyDemoLockOnOpen' in line:
        print(f"Line {i+1}: {line.strip()}")
        # print context
        for j in range(max(0, i-5), min(len(lines), i+6)):
            print(f"  {j+1}: {lines[j].strip()}")
        print("-" * 20)
