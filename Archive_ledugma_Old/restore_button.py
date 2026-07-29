import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.278.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Add the Toggle Button creation back, right before the Exit button.
exit_button_code = """' 11 - Exit System
Dim shpExit As Shape
On Error Resume Next
wsMain.Shapes("btnNavExit").Delete
On Error GoTo ERR_HANDLER
Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 95.9, 385.5, 151.1, 35.3)
"""

toggle_button_code = """' 10 - Toggle Layout
Dim shpToggleLayout As Shape
On Error Resume Next
wsMain.Shapes("shpToggleLayout").Delete
On Error GoTo ERR_HANDLER
Set shpToggleLayout = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 96.6, 265.1, 151.1, 35.4)
shpToggleLayout.Name = "shpToggleLayout"
shpToggleLayout.Fill.ForeColor.RGB = RGB(80, 80, 80)
If ActiveWindow.Zoom = 75 Then
    shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1512) & ChrW(1495) & ChrW(1489)
Else
    shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)
End If
shpToggleLayout.TextFrame2.TextRange.Font.Size = 13
shpToggleLayout.TextFrame2.TextRange.Font.Bold = msoTrue
shpToggleLayout.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
shpToggleLayout.OnAction = "ToggleDisplayMode"

"""

# Replace the existing Exit button block with the Toggle + Exit (with Exit moved back to 385.5)
existing_exit = """' 11 - Exit System
Dim shpExit As Shape
On Error Resume Next
wsMain.Shapes("btnNavExit").Delete
On Error GoTo ERR_HANDLER
Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 95.9, 265.1, 151.1, 35.3)"""

content = content.replace(existing_exit, toggle_button_code + exit_button_code)

# 2. Make sure the Auto-Zoom in A00 also sets the button text
auto_zoom_old = """    ' Auto-detect screen width and set zoom
    If Application.UsableWidth < 1200 Then
        ActiveWindow.Zoom = 75
    Else
        ActiveWindow.Zoom = 100
    End If"""

auto_zoom_new = """    ' Auto-detect screen width and set zoom
    If Application.UsableWidth < 1200 Then
        ActiveWindow.Zoom = 75
        wsMain.Shapes("shpToggleLayout").TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1512) & ChrW(1495) & ChrW(1489)
    Else
        ActiveWindow.Zoom = 100
        wsMain.Shapes("shpToggleLayout").TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)
    End If"""

content = content.replace(auto_zoom_old, auto_zoom_new)

# 3. Add the ToggleDisplayMode subroutine at the end
toggle_sub = """

Public Sub ToggleDisplayMode()
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(ChrW(1491) & ChrW(1508) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514))
    
    If wsMain Is Nothing Then Exit Sub
    
    wsMain.Unprotect Password:="1234"
    
    Dim shpToggleLayout As Shape
    Set shpToggleLayout = wsMain.Shapes("shpToggleLayout")
    
    If ActiveWindow.Zoom > 80 Then
        ActiveWindow.Zoom = 75
        If Not shpToggleLayout Is Nothing Then
            shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1512) & ChrW(1495) & ChrW(1489)
        End If
    Else
        ActiveWindow.Zoom = 100
        If Not shpToggleLayout Is Nothing Then
            shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)
        End If
    End If
    
    wsMain.Protect Password:="1234"
End Sub
"""
content = content + toggle_sub

# Update versions
content = content.replace('APP_VERSION As String = "2.278"', 'APP_VERSION As String = "2.279"')
content = content.replace('VERSION: V2.278', 'VERSION: V2.279')
content = content.replace('Error in V2.278!', 'Error in V2.279!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_278"', 'Attribute VB_Name = "Goren_Claude_V2_279"')

changelog = """' CHANGES IN 2.279:
'   - UI: Restored Toggle Layout button but purely to toggle zoom manually (75% / 100%) for users dragging across monitors.
"""
content = content.replace("' CHANGES IN 2.278:", changelog + "' CHANGES IN 2.278:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.279.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 279')
