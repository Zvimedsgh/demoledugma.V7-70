import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.128.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target = """wsSearch.Range("E5:E9").Font.Size = 12
wsSearch.Range("E5:E9").Font.Color = RGB(0, 112, 192)
wsSearch.Range("E5:E9").HorizontalAlignment = -4152 ' xlRight
wsSearch.Cells(5, 5).Font.Bold = True"""

new_target = """wsSearch.Range("E5:E9").Font.Size = 12
wsSearch.Range("E5:E9").Font.Color = RGB(0, 112, 192)
wsSearch.Cells(5, 5).Font.Bold = True"""

if target in content:
    content = content.replace(target, new_target)
    print("Replaced horizontal alignment.")
else:
    print("Not found horizontal alignment.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_128"', 'Attribute VB_Name = "Goren_Claude_V2_129"')
content = content.replace('VERSION: V2.128', 'VERSION: V2.129')
content = content.replace('APP_VERSION As String = "2.128"', 'APP_VERSION As String = "2.129"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.129.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.129 created.")
