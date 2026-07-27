import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.199.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(4470, 4495):
    print(f"[{i+1}] {lines[i].strip()}")

print("---")
for i in range(4398, 4408):
    print(f"[{i+1}] {lines[i].strip()}")
