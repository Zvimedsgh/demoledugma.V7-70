with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'Sub BuildHomePage\b.*?(?=End Sub)', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(0)
    for i, line in enumerate(sub_text.split('\n')[:30]):
        print(f"{i+1}: {line}")
