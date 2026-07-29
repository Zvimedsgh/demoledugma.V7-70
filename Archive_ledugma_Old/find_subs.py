with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
matches = re.finditer(r'^(?:Public\s+|Private\s+)?Sub\s+([A-Za-z0-9_]+)', text, re.MULTILINE)
for m in matches:
    if 'Setup' in m.group(1) or 'Build' in m.group(1) or 'Admin' in m.group(1):
        print(m.group(1))
