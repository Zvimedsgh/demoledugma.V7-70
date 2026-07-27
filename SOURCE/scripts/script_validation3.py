import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.017.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

import re
# Replace Validation.Delete on merged ranges G5 to G10
content = re.sub(r'wsMain\.Range\("G([5-9]|10)"\)\.Validation\.Delete', r'wsMain.Range("G\1").MergeArea.Validation.Delete', content)
content = re.sub(r'wsMain\.Range\("G([5-9]|10)"\)\.Validation\.Add', r'wsMain.Range("G\1").MergeArea.Validation.Add', content)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
