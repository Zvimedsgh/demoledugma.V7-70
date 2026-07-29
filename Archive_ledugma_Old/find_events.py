with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
matches = re.finditer(r'(?:Private\s+|Public\s+)?(?:Sub|Function)\s+([A-Za-z0-9_]+)', text)
for m in matches:
    name = m.group(1)
    if 'event' in name.lower() or 'change' in name.lower() or 'activate' in name.lower() or 'open' in name.lower():
        print(name)
