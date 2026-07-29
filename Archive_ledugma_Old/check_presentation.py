with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Sub BuildPresentation\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    for i, line in enumerate(sub_text.split('\n')):
        if '.Visible' in line or 'Unprotect' in line or 'Protect' in line or 'Error' in line:
            print(f"{i+1}: {line.strip()}")
