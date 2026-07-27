import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.094.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.094.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('wsMain.Range("K4").Value = 3.9', 'wsMain.Range("K4").Value = 3.4239')

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

