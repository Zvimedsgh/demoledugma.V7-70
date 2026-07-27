import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

print("Line 2044 context:")
for i in range(2040, 2048):
    print(lines[i].strip())

print("\nLine 2505 context:")
for i in range(2500, 2508):
    print(lines[i].strip())

print("\nLine 6950 context:")
for i in range(6945, 6955):
    print(lines[i].strip())

