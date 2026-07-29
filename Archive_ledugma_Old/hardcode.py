import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.272.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace Operation Manual
old_op = 'Set shpOpManual = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C16").Top, btnW, btnH)'
new_op = 'Set shpOpManual = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 96.6, 171.3, 151.1, 35.3)'
content = content.replace(old_op, new_op)

# Replace Toggle Sheets
old_ts = 'Set shpToggle = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C18").Top, btnW, btnH)'
new_ts = 'Set shpToggle = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 98.1, 218.2, 151.1, 35.4)'
content = content.replace(old_ts, new_ts)

# Replace Toggle Layout
old_tl = 'Set shpToggleLayout = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C20").Top, btnW, btnH)'
new_tl = 'Set shpToggleLayout = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 96.6, 265.1, 151.1, 35.4)'
content = content.replace(old_tl, new_tl)

# Replace Exit
old_exit = 'Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C22").Top, btnW, btnH)'
new_exit = 'Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 95.9, 385.5, 151.1, 35.3)'
content = content.replace(old_exit, new_exit)

content = content.replace('APP_VERSION As String = "2.272"', 'APP_VERSION As String = "2.273"')
content = content.replace('VERSION: V2.272', 'VERSION: V2.273')
content = content.replace('Error in V2.272!', 'Error in V2.273!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_272"', 'Attribute VB_Name = "Goren_Claude_V2_273"')

changelog = """' CHANGES IN 2.273:
'   - UI: Hardcoded exact custom coordinates for system buttons in Desktop Mode as visually arranged by the user.
"""
content = content.replace("' CHANGES IN 2.272:", changelog + "' CHANGES IN 2.272:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.273.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 273')
