with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

import re

for i, line in enumerate(lines):
    l = line.strip()
    if not l or l.startswith("'") or l.startswith("Rem ") or l.startswith("#"):
        continue
        
    # Check if line looks like an english sentence (has no '=' or '.' and starts with capital letter but isn't a known keyword)
    if re.match(r'^[A-Z][a-z]+ [a-z]+ ', l) and '=' not in l and '.' not in l and '(' not in l and not l.startswith('On Error') and not l.startswith('End ') and not l.startswith('Exit '):
        print(f"Line {i+1}: {l}")
        
    # Specifically look for 'If empty,'
    if l.startswith("If empty,"):
        print(f"Line {i+1}: {l}")
