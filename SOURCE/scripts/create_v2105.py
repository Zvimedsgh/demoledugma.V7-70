import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.104.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Add Auto_Open at the top of the module, right after the constants
content = content.replace('Public Sub NavToIndex()', '''Public Sub Auto_Open()
    ' Ensure Home page is set up correctly when the file opens
    On Error Resume Next
    Call A00_SetupMainSheet
End Sub

Public Sub NavToIndex()''')

# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_104"', 'Attribute VB_Name = "Goren_Claude_V2_105"')
content = content.replace('VERSION: V2.104', 'VERSION: V2.105')
content = content.replace('APP_VERSION As String = "2.104"', 'APP_VERSION As String = "2.105"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.105.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.105 created with Auto_Open.")
