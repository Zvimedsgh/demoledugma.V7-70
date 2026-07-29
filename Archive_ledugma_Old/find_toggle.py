with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
matches = re.finditer(r'(?:Sub|Function)\s+([A-Za-z0-9_]+)', text)
for m in matches:
    name = m.group(1)
    if 'hide' in name.lower() or 'show' in name.lower() or 'toggle' in name.lower():
        print(name)
