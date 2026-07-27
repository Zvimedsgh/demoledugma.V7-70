import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

lines = content.splitlines()
for i, line in enumerate(lines):
    if "Public Sub SaveReportsToFolder()" in line:
        print(f"Line {i+1}: {line}")
        print(f"Line {i+2}: {lines[i+1]}")
        print(f"Line {i+3}: {lines[i+2]}")
    if "Public Sub BuildReview()" in line:
        print(f"Line {i+1}: {line}")
        print(f"Line {i+2}: {lines[i+1]}")
        print(f"Line {i+3}: {lines[i+2]}")
