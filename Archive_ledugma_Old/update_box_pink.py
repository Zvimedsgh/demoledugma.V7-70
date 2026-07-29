with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
pattern = r'(shpInstall\.Fill\.ForeColor\.RGB = )RGB\(\d+, \d+, \d+\).*'
replacement = r'\1RGB(255, 153, 153) \' Prominent pastel pink'

text = re.sub(pattern, replacement, text)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated shape background to pink successfully.")
