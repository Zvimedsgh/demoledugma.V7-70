with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
for i, line in enumerate(text.split('\n')):
    if 'הוראות התקנה' in line and 'AddShape' not in line:
        pass
    if 'AddShape' in line or 'btnNav' in line:
        pass

match = re.search(r'(Sub SetupAdminProtection\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    for i, line in enumerate(sub_text.split('\n')):
        if 'TextRange.Text' in line and 'הוראות' in line:
            print(f"{i+1}: {line.strip()}")
            print(f"Shape name: {sub_text.splitlines()[i-2].strip()}")
