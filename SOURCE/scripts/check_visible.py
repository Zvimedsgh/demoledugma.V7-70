import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.101.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

matches = re.findall(r'.{0,50}Visible =.{0,50}', content, re.IGNORECASE)
for m in matches:
    if "SetupMain" in m or "xlSheet" in m:
        pass
    print(m)

