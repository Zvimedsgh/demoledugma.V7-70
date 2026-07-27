import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.055.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(1010, 1020):
    print(f"{i+1}: {lines[i].strip()}")

print("-" * 20)

for i in range(1450, 1460):
    print(f"{i+1}: {lines[i].strip()}")
