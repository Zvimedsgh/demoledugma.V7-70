with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

pattern = r'    On Error Resume Next\n    wsMain\.Shapes\("shpInstallMsg"\)\.Delete\n    On Error GoTo 0\n    \n    Dim shpInstall As Shape\n    \' Place it in A2, adjust size\n    Set shpInstall = wsMain\.Shapes\.AddShape\(msoShapeRectangle, wsMain\.Range\("A2"\)\.Left \+ 10, wsMain\.Range\("A2"\)\.Top, 220, 65\)\n    shpInstall\.Name = "shpInstallMsg"'

replacement = """    Dim shpInstall As Shape
    On Error Resume Next
    Set shpInstall = wsMain.Shapes("shpInstallMsg")
    On Error GoTo 0
    If shpInstall Is Nothing Then
        Set shpInstall = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("E16").Left, wsMain.Range("E16").Top, 220, 65)
        shpInstall.Name = "shpInstallMsg"
    End If"""

text = re.sub(pattern, replacement, text, flags=re.DOTALL)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated shape creation logic.")
