with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    l = line.strip()
    if not l or l.startswith("'") or l.startswith("Rem ") or l.startswith("#"):
        continue
    
    # Let's find any line that does not start with a valid VBA keyword or object
    # We will just print any line that contains '|' and isn't a string assignment
    if '|' in l and '"|"' not in l and '=|' not in l and 'ChrW' not in l and '& "|"' not in l:
        print(f"Line {i+1}: {l}")
        
    # Also find 'insuredRef' at the start of a line
    if l.startswith("insuredRef"):
        print(f"Line {i+1}: {l}")
