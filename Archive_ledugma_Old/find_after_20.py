with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Sub A00_SetupMainSheet\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    found_20 = False
    for line in sub_text.split('\n'):
        if line.strip().startswith('20 '):
            found_20 = True
        if found_20:
            print(line.rstrip())
            if line.strip().startswith('30 '):
                break
