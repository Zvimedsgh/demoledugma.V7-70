with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas','r',encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'Sub BuildReview\b.*?(?=End Sub)', text, re.DOTALL | re.IGNORECASE)
if match:
    code = match.group(0)
    lines = code.split('\n')
    for i in range(max(0, 600-10), min(len(lines), 600+10)):
        print(f"{i}: {lines[i]}")
