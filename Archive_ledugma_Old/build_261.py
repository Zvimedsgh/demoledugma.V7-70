import os
import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.260.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.260"', 'APP_VERSION As String = "2.261"')
content = content.replace('VERSION: V2.260', 'VERSION: V2.261')
content = content.replace('Error in V2.260!', 'Error in V2.261!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_260"', 'Attribute VB_Name = "Goren_Claude_V2_261"')

# Update Changelog
changelog = """' CHANGES IN 2.261:
'   - BUGFIX: Fixed Error 1004 on G3:G4 Interior Color by addressing cells individually with MergeArea.
"""
content = content.replace("' CHANGES IN 2.260:", changelog + "' CHANGES IN 2.260:")

# Fix line 1740
old_1740 = '1740 wsMain.Range("G3:G4").Interior.Color = RGB(230, 230, 230) \' Gray out visually'
new_1740 = '''1740 On Error Resume Next
1741 wsMain.Range("G3").MergeArea.Interior.Color = RGB(230, 230, 230) ' Gray out visually
1742 wsMain.Range("G4").MergeArea.Interior.Color = RGB(230, 230, 230)'''
content = content.replace(old_1740, new_1740)

# Fix line 1760 Validation
old_1760 = '''1760 With wsMain.Range("G3:G4").Validation
1770     .Delete
1780 End With'''
new_1760 = '''1760 On Error Resume Next
1765 wsMain.Range("G3").Validation.Delete
1770 wsMain.Range("G4").Validation.Delete
1780 Err.Clear'''
content = content.replace(old_1760, new_1760)

# Fix line 1820
old_1820 = '1820 wsMain.Range("G3:G4").Interior.Color = xlNone'
new_1820 = '''1820 On Error Resume Next
1821 wsMain.Range("G3").MergeArea.Interior.Color = xlNone
1822 wsMain.Range("G4").MergeArea.Interior.Color = xlNone'''
content = content.replace(old_1820, new_1820)

# Fix lock rectangle (doesn't usually fail but just to be safe)
# Set shpLock = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("G3:G4").Left, wsMain.Range("G3:G4").Top, wsMain.Range("G3:G4").Width, wsMain.Range("G3:G4").Height)
old_shape = 'Set shpLock = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("G3:G4").Left, wsMain.Range("G3:G4").Top, wsMain.Range("G3:G4").Width, wsMain.Range("G3:G4").Height)'
new_shape = 'Set shpLock = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("G3").Left, wsMain.Range("G3").Top, wsMain.Range("G3").Width, wsMain.Range("G3").Height + wsMain.Range("G4").Height)'
content = content.replace(old_shape, new_shape)

# Fix the grey color index
old_grey = 'wsMain.Range("G3:G4").Interior.ColorIndex = 15 \' Grey'
new_grey = '''wsMain.Range("G3").MergeArea.Interior.ColorIndex = 15
wsMain.Range("G4").MergeArea.Interior.ColorIndex = 15'''
content = content.replace(old_grey, new_grey)

# Fix another validation delete
old_val_del = 'wsMain.Range("G3:G4").Validation.Delete'
new_val_del = '''wsMain.Range("G3").Validation.Delete
wsMain.Range("G4").Validation.Delete'''
content = content.replace(old_val_del, new_val_del)

# Fix validation delete again
old_val_del2 = 'With wsMain.Range("G3:G4").Validation'
new_val_del2 = '''With wsMain.Range("G3").Validation
.Delete
End With
With wsMain.Range("G4").Validation'''
content = content.replace(old_val_del2, new_val_del2)

# Fix the other Interior Color
old_color = 'wsMain.Range("G3:G4").Interior.Color = RGB(255, 245, 230)'
new_color = '''wsMain.Range("G3").MergeArea.Interior.Color = RGB(255, 245, 230)
wsMain.Range("G4").MergeArea.Interior.Color = RGB(255, 245, 230)'''
content = content.replace(old_color, new_color)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.261.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 261')
