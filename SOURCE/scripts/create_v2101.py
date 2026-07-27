import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.100.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.101.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update all V2.095 to V2.101 in the header!
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_095"', 'Attribute VB_Name = "Goren_Claude_V2_101"')
content = content.replace('VERSION: V2.095', 'VERSION: V2.101')
content = content.replace('Private Const APP_VERSION As String = "2.095"', 'Private Const APP_VERSION As String = "2.101"')
# And also replace any leftover V2.099 or V2.100 just in case
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_100"', 'Attribute VB_Name = "Goren_Claude_V2_101"')
content = content.replace('VERSION: V2.100', 'VERSION: V2.101')
content = content.replace('Private Const APP_VERSION As String = "2.100"', 'Private Const APP_VERSION As String = "2.101"')

# 2. Fix GetStringParameter to remove EnableEvents
def fix_getstring(match):
    return match.group(0).replace("Application.EnableEvents = False\n", "").replace("Application.EnableEvents = True\n", "")
content = re.sub(r'Private Function GetStringParameter\(.*?(?=End Function)End Function', fix_getstring, content, flags=re.DOTALL)

# 3. Update NavToIndex and NavSettings_Home to call A00_SetupMainSheet
old_nav_index = """Public Sub NavToIndex()
    Call HideWorkSheets
End Sub"""
new_nav_index = """Public Sub NavToIndex()
    Call A00_SetupMainSheet
    Call HideWorkSheets
    On Error Resume Next
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("F3").Select
End Sub"""
content = content.replace(old_nav_index, new_nav_index)

old_nav_home = """Public Sub NavSettings_Home()
    Call HideWorkSheets
End Sub"""
new_nav_home = """Public Sub NavSettings_Home()
    Call A00_SetupMainSheet
    Call HideWorkSheets
    On Error Resume Next
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("F3").Select
End Sub"""
content = content.replace(old_nav_home, new_nav_home)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.101 created.")
