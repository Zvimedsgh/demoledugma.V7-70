with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
subs = ['OpenManual', 'OpenInstructions']

for sub in subs:
    match = re.search(fr'(Sub {sub}\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
    if match:
        print(f"--- {sub} ---")
        for i, line in enumerate(match.group(1).split('\n')):
            print(f"{i+1}: {line.strip()}")
