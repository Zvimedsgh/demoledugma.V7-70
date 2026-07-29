with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
matches = re.finditer(r'(Public Sub [A-Za-z0-9_]+|Sub [A-Za-z0-9_]+)', text, re.IGNORECASE)
for m in matches:
    name = m.group(1)
    if 'manual' in name.lower() or 'inst' in name.lower() or 'open' in name.lower():
        print(name)
