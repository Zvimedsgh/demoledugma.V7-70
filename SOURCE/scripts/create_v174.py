import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.173.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.174.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    lines = f.readlines()

old_delete = """    ' ---- Remove ALL old buttons ----
    2790 On Error Resume Next
    2800 For Each s In wsMain.Shapes
        2810 s.Delete
    2820 Next s
    2830 Err.Clear"""

new_delete = """    ' ---- Remove ALL old buttons and stubborn floating shapes ----
    2790 On Error Resume Next
    2800 wsMain.DrawingObjects.Delete ' Force delete ALL drawing objects!
    2810 For Each s In wsMain.Shapes
        2815 s.Delete ' Fallback for anything DrawingObjects missed
    2820 Next s
    2830 Err.Clear"""

content = "".join(lines)
content = content.replace(old_delete, new_delete)

# Update version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_173"', 'Attribute VB_Name = "Goren_Claude_V2_174"')
content = content.replace('VERSION: V2.173', 'VERSION: V2.174')
content = content.replace('APP_VERSION As String = "2.173"', 'APP_VERSION As String = "2.174"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.174 correctly!")
