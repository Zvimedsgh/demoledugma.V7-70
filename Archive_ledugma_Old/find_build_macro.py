with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
matches = re.finditer(r'Sub\s+([A-Za-z0-9_]+)\(\)', text)
for m in matches:
    sub_name = m.group(1)
    # Check if this sub contains the WhatsApp string or the A22 version string
    sub_text = text[m.start():m.start()+5000] # roughly
    if '054-6677396' in sub_text or 'wsMain.Range("A22")' in sub_text:
        print(sub_name)
