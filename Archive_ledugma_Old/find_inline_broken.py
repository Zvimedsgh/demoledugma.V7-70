import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    l = line.strip()
    if not l or l.startswith("'") or l.startswith("Rem "):
        continue
    if re.search(r'\b(This is|Note:|TODO|FIXME|BUG|HACK|MODULE:|PURPOSE:|VERSION:)\b', l, re.IGNORECASE) and not l.startswith('"'):
        print(f"Line {i+1}: {l}")
