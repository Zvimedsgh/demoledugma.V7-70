import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_str = """    If GetActiveAgencyName = "" Then
        GetActiveAgencyName = GetActiveAgencyName() ' "Levav System" default
    End If"""

new_str = """    If GetActiveAgencyName = "" Then
        GetActiveAgencyName = "Levav System"
    End If"""

if old_str in content:
    content = content.replace(old_str, new_str)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Fixed GetActiveAgencyName bug")
else:
    print("Could not find GetActiveAgencyName string")
