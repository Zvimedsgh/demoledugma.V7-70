with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Sub ProcessDataWrapper\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    print(sub_text[:500])
