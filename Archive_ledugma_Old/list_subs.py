with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
matches = re.finditer(r'(Public Sub [A-Za-z0-9_]+)', text, re.IGNORECASE)
for m in matches:
    print(m.group(1))
