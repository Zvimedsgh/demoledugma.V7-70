import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.028.bas', 'r', encoding='utf-8') as f:
    content = f.read()
    m = re.search(r'Public Sub NavSettings_FieldMap\(\).*?End Sub', content, re.DOTALL)
    if m:
        print(m.group(0))
