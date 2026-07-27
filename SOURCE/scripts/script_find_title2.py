import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

lines = content.splitlines()
for i, line in enumerate(lines):
    if "TextFrame" in line and "ChrW(1500) & ChrW(1489) & ChrW(1489)" in line:
        print(f"Line {i+1}: {line}")
