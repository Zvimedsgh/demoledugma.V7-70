import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.176.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3840, 3915):
    if "ParagraphFormat.Alignment" in lines[i]:
        print(f"[{i+1}] {lines[i].strip()}")

