import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.181.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

print("Around 5907:")
for i in range(5895, 5910):
    print(f"[{i+1}] {lines[i].strip()}")

print("\nAround 7596:")
for i in range(7585, 7600):
    print(f"[{i+1}] {lines[i].strip()}")

