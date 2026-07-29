with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
matches = re.finditer(r'(?:Private\s+|Public\s+)?(?:Sub|Function)\s+([A-Za-z0-9_]+)', text)
print("List of all Subroutines/Functions:")
subs = []
for m in matches:
    subs.append(m.group(1))

# Check specifically for event handlers (usually in ThisWorkbook or Worksheets, but might be explicitly mapped here)
for s in subs:
    if 'Open' in s or 'Activate' in s or 'Change' in s or 'Event' in s:
        print(f"Potential Event: {s}")

# Check where RGB(0, 150, 0) is used:
print("\nUsages of RGB(0, 150, 0):")
for i, line in enumerate(text.split('\n')):
    if '150, 0)' in line or '150,0)' in line:
        print(f"Line {i+1}: {line.strip()}")
        
# Check where Protect/Unprotect is used:
print("\nProtection logic:")
for i, line in enumerate(text.split('\n')):
    if 'Protect ' in line or 'Protect"' in line or 'ProtectStructure' in line:
        if 'Unprotect' not in line:
            print(f"Line {i+1}: {line.strip()}")
