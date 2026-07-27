import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.020.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_version = '''Private Const APP_VERSION As String = "2.019"'''
new_version = '''Private Const APP_VERSION As String = "2.020"'''

old_vbname = '''Attribute VB_Name = "Goren_Claude_V2_019"'''
new_vbname = '''Attribute VB_Name = "Goren_Claude_V2_020"'''

content = content.replace(old_version, new_version)
content = content.replace(old_vbname, new_vbname)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.020 successfully")
