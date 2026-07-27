import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.101.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.102.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_101"', 'Attribute VB_Name = "Goren_Claude_V2_102"')
content = content.replace('VERSION: V2.101', 'VERSION: V2.102')
content = content.replace('Private Const APP_VERSION As String = "2.101"', 'Private Const APP_VERSION As String = "2.102"')

# Replace the unhide loop in CheckUserPermissions
old_code = """If accessLevel = fullAccess Or UCase$(accessLevel) = "FULLACCESS" Then
' Full access - show all sheets, restore all buttons
For Each ws In ThisWorkbook.Worksheets
ws.Visible = xlSheetVisible
Next ws"""

new_code = """If accessLevel = fullAccess Or UCase$(accessLevel) = "FULLACCESS" Then
' Full access - restore all buttons and hide background sheets
Call HideWorkSheets"""

content = content.replace(old_code, new_code)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.102 created.")
