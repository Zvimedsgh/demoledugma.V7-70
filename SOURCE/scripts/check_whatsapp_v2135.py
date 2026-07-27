import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.135.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

found = False
for i, line in enumerate(lines):
    if "Whatsapp" in line:
        found = True
        print(f"[{i+1}] {lines[i].strip()}")

print(f"Found whatsapp text: {found}")
