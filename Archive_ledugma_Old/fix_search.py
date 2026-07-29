with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# Add Unprotect to SearchClientName
text = re.sub(
    r'(Public Sub SearchClientName\(\)\n\s*On Error GoTo ERR_HANDLER\n)',
    r'\1    On Error Resume Next: ThisWorkbook.Unprotect "Z961814r": On Error GoTo ERR_HANDLER\n',
    text,
    flags=re.IGNORECASE
)

# Add Unprotect to DoClientSearch
text = re.sub(
    r'(Public Sub DoClientSearch\(\)\n\s*On Error GoTo ERR_HANDLER\n)',
    r'\1    On Error Resume Next: ThisWorkbook.Unprotect "Z961814r": On Error GoTo ERR_HANDLER\n',
    text,
    flags=re.IGNORECASE
)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Fixed SearchClientName and DoClientSearch!")
