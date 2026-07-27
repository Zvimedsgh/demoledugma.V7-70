import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.100.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    for i in range(20):
        print(f.readline().strip())

