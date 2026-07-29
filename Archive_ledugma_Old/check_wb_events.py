with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
matches = re.finditer(r'(?:Private\s+|Public\s+)?(?:Sub|Function)\s+Workbook_([A-Za-z0-9_]+)', text)
for m in matches:
    print(m.group(1))
