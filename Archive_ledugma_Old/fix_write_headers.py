with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# Remove Protect from WriteComparisonHeaders
text = re.sub(r'490\s+ThisWorkbook\.Protect\s+"Z961814r"\s*\n', '\n', text, flags=re.IGNORECASE)

# Let's search the whole file for other ThisWorkbook.Protect calls and list where they are
matches = re.finditer(r'(?:Sub|Function)\s+([A-Za-z0-9_]+)[^\n]*\n(.*?)End (?:Sub|Function)', text, re.DOTALL | re.IGNORECASE)
for m in matches:
    name = m.group(1)
    sub_text = m.group(2)
    if 'Protect' in sub_text:
        print(f"--- {name} ---")
        for i, line in enumerate(sub_text.split('\n')):
            if 'Protect' in line and 'ThisWorkbook' in line:
                print(f"  {line.strip()}")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Removed Protect from WriteComparisonHeaders!")
