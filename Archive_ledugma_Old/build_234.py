import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.233.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.233', 'VERSION: V2.234')
content = content.replace('Error in V2.233!', 'Error in V2.234!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_233"', 'Attribute VB_Name = "Goren_Claude_V2_234"')
content = content.replace('APP_VERSION As String = "2.233"', 'APP_VERSION As String = "2.234"')

# Force unlock G3:G12 before protecting
unlock_str = '    wsMain.Unprotect "Z961814r"\n    wsMain.Range("G3:G12").Locked = False\n    wsMain.Protect Password:="Z961814r", UserInterfaceOnly:=True'
content = content.replace('    wsMain.Unprotect "Z961814r"\n    wsMain.Protect Password:="Z961814r", UserInterfaceOnly:=True', unlock_str)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.234.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 234')
