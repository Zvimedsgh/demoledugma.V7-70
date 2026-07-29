with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# We will wrap line 1745
pattern = r'(2740\s+BuildSummarySheet[^\n]+)\n(ThisWorkbook\.Worksheets\(SHEET_SUMMARY\(\)\)\.Tab\.Color = RGB\(255, 204, 153\)[^\n]+)'
replacement = r'\1\nOn Error Resume Next\nThisWorkbook.Unprotect "Z961814r"\n\2\nThisWorkbook.Protect "Z961814r"\nOn Error GoTo ERR_HANDLER'

text = re.sub(pattern, replacement, text)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Fixed tab color protection issue!")
