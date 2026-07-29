with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    l = line.strip()
    
    # If we reach the first real code block, stop checking
    if l.startswith('Public Const') or l.startswith('Const ') or l.startswith('Public Type') or l.startswith('Dim ') or l.startswith('Public ') or l.startswith('Private ') or l.startswith('Sub ') or l.startswith('Function '):
        break
        
    if not l:
        continue
        
    # Valid declarations at the top of a VBA module
    if l.startswith("'") or l.startswith('Attribute ') or l.startswith('Option ') or l.startswith('Rem '):
        continue
        
    # Anything else is a broken comment
    print(f"Fixing line {i+1}: {l}")
    # Restore the comment marker
    if not line.lstrip().startswith("'"):
        lines[i] = "' " + line.lstrip()

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.writelines(lines)
print("Fixed remaining broken comments!")
