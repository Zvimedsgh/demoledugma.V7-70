import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.175.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(len(lines)-20, len(lines)):
    print(lines[i].strip())
