with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Sub [A-Za-z0-9_]+).*?3975 wsMain.Range\("A22"\)', text, re.DOTALL)
if match:
    print(match.group(1))
