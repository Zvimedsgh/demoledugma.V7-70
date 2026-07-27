import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.147.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target = r'wsMain.Range(\"A1\").Font.Color = RGB(200, 0, 0) ' + "' Always Red"
new_code = r'wsMain.Range("A1").Font.Color = RGB(200, 0, 0) ' + "' Always Red"

if target in content:
    content = content.replace(target, new_code)
    print("Fixed backslash syntax error!")
else:
    print("Could not find the target string.")
    # let's try a broader replacement just in case
    content = content.replace(r'\"A1\"', '"A1"')

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_147"', 'Attribute VB_Name = "Goren_Claude_V2_148"')
content = content.replace('VERSION: V2.147', 'VERSION: V2.148')
content = content.replace('APP_VERSION As String = "2.147"', 'APP_VERSION As String = "2.148"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.148.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.148 created.")
