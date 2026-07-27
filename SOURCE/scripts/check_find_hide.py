import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.215.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsName1 = ChrW(1492) & ChrW(1493) & ChrW(1512)" in line and "wsName2 = " in lines[i+1]:
        for j in range(i-2, i+5):
            print(lines[j].rstrip())
        break
