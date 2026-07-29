with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
# check BuildBaseSheet
match = re.search(r'(Sub BuildBaseSheet\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match and 'Protect' in match.group(1):
    print("BuildBaseSheet has Protect!")
