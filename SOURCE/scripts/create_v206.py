import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.205.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.206.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

old_block = """' If user not found in list, default to limited
If Not found Then accessLevel = limitedAccess

' Apply permissions
Application.ScreenUpdating = False"""

new_block = """' If user not found in list, default to limited
If Not found Then accessLevel = limitedAccess

' --- OVERRIDE FOR DEMO MODE ---
' In Demo Mode, always grant Full Access so prospective clients can test the buttons!
Dim isDemoPerms As Boolean
Dim demoParamPerms As String
isDemoPerms = FORCE_DEMO_MODE
On Error Resume Next
demoParamPerms = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
On Error GoTo PERM_ERR
If Not FORCE_DEMO_MODE Then
    If demoParamPerms = ChrW(1499) & ChrW(1503) Or demoParamPerms = "YES" Then isDemoPerms = True
End If

If isDemoPerms Then
    accessLevel = fullAccess
End If
' ------------------------------

' Apply permissions
Application.ScreenUpdating = False"""

content = content.replace(old_block, new_block)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_205"', 'Attribute VB_Name = "Goren_Claude_V2_206"')
content = content.replace('VERSION: V2.205', 'VERSION: V2.206')
content = content.replace('APP_VERSION As String = "2.205"', 'APP_VERSION As String = "2.206"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.206")
