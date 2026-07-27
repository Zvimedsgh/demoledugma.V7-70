import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.196.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "1505" in line and "1497" in line and "1504" in line and "1493" in line and "1503" in line:
        for j in range(max(0, i-2), min(len(lines), i+3)):
            print(f"[{j+1}] {lines[j].strip()}")
        print("---")

