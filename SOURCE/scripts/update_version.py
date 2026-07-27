import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_version = '''Private Const APP_VERSION As String = "2.016"'''
new_version = '''Private Const APP_VERSION As String = "2.019"'''

old_date = '''Private Const APP_DATE As String = "01/07/2026 16:01"'''
new_date = '''Private Const APP_DATE As String = "02/07/2026 11:00"'''

old_vbname = '''Attribute VB_Name = "Goren_Claude_V2_018"'''
new_vbname = '''Attribute VB_Name = "Goren_Claude_V2_019"'''

content = content.replace(old_version, new_version)
content = content.replace(old_date, new_date)
content = content.replace(old_vbname, new_vbname)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Updated version headers successfully")
