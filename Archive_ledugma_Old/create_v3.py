import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.279.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Inject the button creation before ' 11 - Exit System
existing_exit = """    ' 11 - Exit System
    Dim shpExit As Shape
    On Error Resume Next
    wsMain.Shapes("btnNavExit").Delete
    On Error GoTo ERR_HANDLER
    Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 95.9, 265.1, 151.1, 35.3)"""

toggle_button_code = """    ' 10 - Toggle Layout
    Dim shpToggleLayout As Shape
    On Error Resume Next
    wsMain.Shapes("shpToggleLayout").Delete
    On Error GoTo ERR_HANDLER
    Set shpToggleLayout = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 96.6, 265.1, 151.1, 35.4)
    shpToggleLayout.Name = "shpToggleLayout"
    shpToggleLayout.Fill.ForeColor.RGB = RGB(80, 80, 80)
    shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1514) & ChrW(1488) & ChrW(1502) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498)
    shpToggleLayout.TextFrame2.TextRange.Font.Size = 13
    shpToggleLayout.TextFrame2.TextRange.Font.Bold = msoTrue
    shpToggleLayout.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shpToggleLayout.OnAction = "ToggleDisplayMode"

    ' 11 - Exit System
    Dim shpExit As Shape
    On Error Resume Next
    wsMain.Shapes("btnNavExit").Delete
    On Error GoTo ERR_HANDLER
    Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 95.9, 385.5, 151.1, 35.3)"""

content = content.replace(existing_exit, toggle_button_code)

# 2. Update Auto-Zoom in A00 (change 75 to 80, remove text changes)
auto_zoom_old = """    ' Auto-detect screen width and set zoom
    If Application.UsableWidth < 1200 Then
        ActiveWindow.Zoom = 75
    Else
        ActiveWindow.Zoom = 100
    End If"""

auto_zoom_new = """    ' Auto-detect screen width and set zoom
    If Application.UsableWidth < 1200 Then
        ActiveWindow.Zoom = 80
    Else
        ActiveWindow.Zoom = 100
    End If"""
content = content.replace(auto_zoom_old, auto_zoom_new)

# 3. Update ToggleDisplayMode to toggle between 80 and 100
toggle_old = """
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

toggle_new = """
Public Sub ToggleDisplayMode()
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(ChrW(1491) & ChrW(1508) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514))
    
    If wsMain Is Nothing Then Exit Sub
    
    wsMain.Unprotect Password:="1234"
    
    If ActiveWindow.Zoom > 85 Then
        ActiveWindow.Zoom = 80
    Else
        ActiveWindow.Zoom = 100
    End If
    
    wsMain.Protect Password:="1234"
End Sub
"""
content = content.replace(toggle_old.strip(), toggle_new.strip())

# Version Bump to V3.00!
content = content.replace('APP_VERSION As String = "2.279"', 'APP_VERSION As String = "3.00"')
content = content.replace('VERSION: V2.279', 'VERSION: V3.00')
content = content.replace('Error in V2.279!', 'Error in V3.00!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_279"', 'Attribute VB_Name = "Goren_Claude_V3_00"')

changelog = """' CHANGES IN 3.00:
'   - V3.00 Promotion!
'   - UI: Fixed missing toggle button (indentation issue).
'   - UI: Added "התאמת מסך" button that cleanly toggles between 100% and 80%.
"""
content = content.replace("' CHANGES IN 2.279:", changelog + "' CHANGES IN 2.279:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.00.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.00')
