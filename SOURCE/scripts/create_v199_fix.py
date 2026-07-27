import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.198.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.199.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix UpdateFilterValueDropdown error handling around wsLists modification
pattern = r'On Error Resume Next\s+wsLists\.Unprotect "Z961814r"\s+On Error GoTo ERR_HANDLER\s+wsLists\.Range\("T:T"\)\.ClearContents'
replacement = 'On Error Resume Next\n            wsLists.Unprotect "Z961814r"\n            wsLists.Range("T:T").ClearContents'
content = re.sub(pattern, replacement, content)

# Fix early filterValue clearing
pattern2 = r'(\s+)(\d+\s+)(wsMain\.Range\("rngFilterValue"\)\.Value = "")(\s+)(\d+\s+)(On Error Resume Next)'
replacement2 = r'\1\2On Error Resume Next\1\5\3\1\5wsMain.Range("rngFilterValue").Validation.Delete'

# Actually just doing it manually via string index since regex with line numbers is annoying
lines = content.split('\n')
for i, line in enumerate(lines):
    if 'wsMain.Range("rngFilterValue").Value = ""' in line:
        lines[i] = '        On Error Resume Next\n' + line
    if 'ThisWorkbook.Names.Add "lst_temp_filter", wsLists.Range(wsLists.Cells(2, 20), wsLists.Cells(rUI, 20))' in line:
        lines[i] = line + '\n        On Error GoTo ERR_HANDLER'

content = '\n'.join(lines)

# Update version strings
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_198"', 'Attribute VB_Name = "Goren_Claude_V2_199"')
content = content.replace('VERSION: V2.198', 'VERSION: V2.199')
content = content.replace('APP_VERSION As String = "2.198"', 'APP_VERSION As String = "2.199"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.199")
