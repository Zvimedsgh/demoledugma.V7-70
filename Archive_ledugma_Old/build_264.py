import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.263.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.263"', 'APP_VERSION As String = "2.264"')
content = content.replace('VERSION: V2.263', 'VERSION: V2.264')
content = content.replace('Error in V2.263!', 'Error in V2.264!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_263"', 'Attribute VB_Name = "Goren_Claude_V2_264"')

# Changelog
changelog = """' CHANGES IN 2.264:
'   - BUGFIX: Added MergeArea.Value to G5, G8, G10 assignments and expanded On Error Resume Next to prevent any further merged cell crashes in the parameter table.
"""
content = content.replace("' CHANGES IN 2.263:", changelog + "' CHANGES IN 2.263:")

# Fix the value assignments
old_1840 = '''    1840 If IsEmpty(wsMain.Range("G5").Value) Then wsMain.Range("G5").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    1850 If IsEmpty(wsMain.Range("G8").Value) Then wsMain.Range("G8").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
    1860 If IsEmpty(wsMain.Range("G10").Value) Then wsMain.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)

    1870 On Error Resume Next'''

new_1840 = '''    1870 On Error Resume Next
    1840 If IsEmpty(wsMain.Range("G5").Value) Then wsMain.Range("G5").MergeArea.Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    1850 If IsEmpty(wsMain.Range("G8").Value) Then wsMain.Range("G8").MergeArea.Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
    1860 If IsEmpty(wsMain.Range("G10").Value) Then wsMain.Range("G10").MergeArea.Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
'''

content = content.replace(old_1840, new_1840)

# Fix 1950, 1960
content = content.replace('1950 wsMain.Range("G7").Value', '1950 wsMain.Range("G7").MergeArea.Value')
content = content.replace('1960 wsMain.Range("G8").Value', '1960 wsMain.Range("G8").MergeArea.Value')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.264.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 264')
