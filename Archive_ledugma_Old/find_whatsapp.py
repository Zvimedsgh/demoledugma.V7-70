with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
for i, line in enumerate(text.split('\n')):
    if '054-6677396' in line or 'WhatsApp' in line or 'F19' in line:
        print(f"{i+1}: {line.strip()}")
