import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.102.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

def repl_fullaccess(match):
    return """If accessLevel = fullAccess Or UCase$(accessLevel) = "FULLACCESS" Then
' Full access - restore all buttons and hide background sheets
Call HideWorkSheets
"""

content = re.sub(r'If accessLevel = fullAccess Or UCase\$\(accessLevel\) = "FULLACCESS" Then\s*\n\s*\' Full access - show all sheets, restore all buttons\s*\n\s*For Each ws In ThisWorkbook\.Worksheets\s*\n\s*ws\.Visible = xlSheetVisible\s*\n\s*Next ws', repl_fullaccess, content, flags=re.DOTALL)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.102 updated.")
