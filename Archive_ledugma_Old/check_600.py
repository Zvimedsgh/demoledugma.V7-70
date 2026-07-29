with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas','r',encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'Sub BuildReview\b.*?(?=End Sub)', text, re.DOTALL | re.IGNORECASE)
if match:
    code = match.group(0)
    for line in code.split('\n'):
        if '600' in line:
            print(line.strip())
