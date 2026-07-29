import sys

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.12.bas', 'r', encoding='utf-8') as f:
    content = f.read()

target1 = 'If demoParamUI = ChrW(1499) & ChrW(1503) Or demoParamUI = "YES" Then isDemoModeUI = True'
replacement1 = 'If demoParamUI = ChrW(1499) & ChrW(1503) Or demoParamUI = "YES" Or demoParamUI = "" Then isDemoModeUI = True'

if target1 in content:
    content = content.replace(target1, replacement1)
    print("Replaced demoParamUI successfully")
else:
    print("Target 1 not found")

target2 = 'If demoParamStr = ChrW(1499) & ChrW(1503) Or demoParamStr = "YES" Then isSysDemo = True'
replacement2 = 'If demoParamStr = ChrW(1499) & ChrW(1503) Or demoParamStr = "YES" Or demoParamStr = "" Then isSysDemo = True'

if target2 in content:
    content = content.replace(target2, replacement2)
    print("Replaced demoParamStr successfully")
else:
    print("Target 2 not found")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.12.bas', 'w', encoding='utf-8') as f:
    f.write(content)
