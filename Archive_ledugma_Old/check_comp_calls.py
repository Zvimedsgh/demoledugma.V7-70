with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Sub BuildComparisonSheet\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    for i, line in enumerate(sub_text.split('\n')):
        if re.match(r'^\s*([A-Za-z0-9_]+)\s', line) and not any(k in line for k in ['If', 'End', 'Next', 'For', 'Dim', 'Set', 'On Error', 'With', 'MsgBox', 'Let']):
            print(f"Line {i+1}: {line.strip()}")
        elif 'Call ' in line:
            print(f"Line {i+1}: {line.strip()}")
