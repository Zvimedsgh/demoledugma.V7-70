import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.197.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.198.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix UpdateFilterValueDropdown wsLists.ClearContents
old_clear = r'(Set wsLists = ThisWorkbook\.Worksheets\(listsName\)\s+)(wsLists\.Range\("T:T"\)\.ClearContents)'
new_clear = r'\1On Error Resume Next\n            wsLists.Unprotect "Z961814r"\n            On Error GoTo ERR_HANDLER\n            \2'
content = re.sub(old_clear, new_clear, content)

# Update version strings
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_197"', 'Attribute VB_Name = "Goren_Claude_V2_198"')
content = content.replace('VERSION: V2.197', 'VERSION: V2.198')
content = content.replace('APP_VERSION As String = "2.197"', 'APP_VERSION As String = "2.198"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.198")
