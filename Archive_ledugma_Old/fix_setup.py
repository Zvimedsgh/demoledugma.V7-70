with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
text = re.sub(r'(4860\s+wsMatachSetup\.Visible\s*=\s*xlSheetVeryHidden)', r'On Error Resume Next\nThisWorkbook.Unprotect "Z961814r"\n\1', text)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Fixed A00_SetupMainSheet crash!")
