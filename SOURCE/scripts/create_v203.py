import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.202.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.203.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix UpdateFilterValueDropdown Demo Mode logic using regex
pattern = r'(If Not wsDataUI Is Nothing Then\s+)Dim dCol As Long(\s+If filterType = H_COMPANY\(\) Then\s+)dCol = BASE_COL_COMPANY(\s+ElseIf filterType = H_TELLER\(\) Then\s+)dCol = BASE_COL_TELLER(\s+ElseIf filterType = H_AGENT\(\) Then\s+)dCol = BASE_COL_AGENTNAME(\s+ElseIf filterType = H_BRANCH\(\) Then\s+)dCol = BASE_COL_BRANCHNAME(\s+ElseIf filterType = H_BRANCH\(\) & ChrW\(32\) & ChrW\(1502\) & ChrW\(1512\) & ChrW\(1499\) & ChrW\(1494\) Then\s+)dCol = BASE_COL_MAINBRANCH(\s+Else)'

replacement = r'\1InitRawColumns\n        Dim dCol As Long\2dCol = RAW_COMPANY\3dCol = RAW_TELLERNAME\4dCol = RAW_AGENTNAME\5dCol = RAW_BRANCHNAME\6dCol = RAW_BRANCHNAME\7'

content = re.sub(pattern, replacement, content)

# Check if it replaced
if 'InitRawColumns\n        Dim dCol As Long' in content:
    print("Replaced successfully!")
else:
    print("REPLACEMENT FAILED!")

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_202"', 'Attribute VB_Name = "Goren_Claude_V2_203"')
content = content.replace('VERSION: V2.202', 'VERSION: V2.203')
content = content.replace('APP_VERSION As String = "2.202"', 'APP_VERSION As String = "2.203"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.203")
