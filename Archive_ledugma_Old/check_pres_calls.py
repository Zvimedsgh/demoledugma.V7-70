with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Sub BuildPresentation\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    for i in range(1, 120):
        line = sub_text.split('\n')[i].strip()
        if re.match(r'^[A-Za-z0-9_]+', line) and not any(k in line for k in ['If ', 'End ', 'Next', 'For ', 'Dim ', 'Set ', 'On Error ', 'With ', 'Let ', 'MsgBox']):
            print(f"Line {i+1}: {line}")
