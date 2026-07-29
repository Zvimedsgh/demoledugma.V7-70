import re

with open(r'C:\ledugma\DEMO\modDemoReports_V9.40.bas', 'r', encoding='windows-1255', errors='ignore') as f:
    lines = f.readlines()

if_stack = []

for i, line in enumerate(lines):
    l = line.strip()
    if l.startswith("'"): continue # skip comments
    
    # check for multi-line If
    # a multi-line If ends with 'Then' (or 'Then \'' but we stripped trailing comments)
    if l.lower().startswith("if "):
        # remove inline comments
        if "'" in l:
            l = l.split("'", 1)[0].strip()
        if l.lower().endswith("then"):
            if_stack.append((i+1, l))
        elif "then " in l.lower() and "_" in l: # line continuation?
            pass # simplified assumption
    elif l.lower().startswith("end if"):
        if if_stack:
            if_stack.pop()
        else:
            print(f"Extra End If at line {i+1}")

for line_num, text in if_stack:
    print(f"Unclosed If at line {line_num}: {text}")

