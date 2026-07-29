with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
lines = text.split('\n')
print("All Tab Color/ColorIndex modifications:")
for i, line in enumerate(lines):
    if '.Tab' in line and ('Color' in line or 'RGB' in line or '=' in line):
        print(f"Line {i+1}: {line.strip()}")
        # Check context above it
        for j in range(max(0, i-3), i):
            if 'Set ' in lines[j] or 'With ' in lines[j]:
                print(f"  Context: {lines[j].strip()}")
