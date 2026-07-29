with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
lines = text.split('\n')
for i, line in enumerate(lines):
    if 'OnAction' in line and 'RunMainProcess' in line or 'Process' in line:
        print(f"Line {i+1}: {line.strip()}")
