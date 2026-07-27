import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.213.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "MsgBoxU" in line and "1492" in line and "1502" in line and "1510" in line:
        for j in range(i-5, i+5):
            print(lines[j].rstrip())
        break
