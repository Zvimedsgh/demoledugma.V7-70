import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.269.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.269"', 'APP_VERSION As String = "2.270"')
content = content.replace('VERSION: V2.269', 'VERSION: V2.270')
content = content.replace('Error in V2.269!', 'Error in V2.270!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_269"', 'Attribute VB_Name = "Goren_Claude_V2_270"')

changelog = """' CHANGES IN 2.270:
'   - UI: Complete layout revamp. "System Tools" (Toggle Sheets, Toggle Layout, Exit, Operation Manual) unified under Main Menu in Column C.
'   - UI: Phone number footer moved to row 24.
"""
content = content.replace("' CHANGES IN 2.269:", changelog + "' CHANGES IN 2.269:")

# --- Modify A00_SetupMainSheet ---

# Remove old ToggleHiddenSheets button creation
old_toggle_sheets = """    3760 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("L19").Left, wsMain.Range("L19").Top, 160, 30)
    3770 shp.Fill.ForeColor.RGB = RGB(80, 80, 80)
    3780 shp.Line.Visible = msoFalse
    3790 shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & "/" & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1512) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514)
    3800 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    3810 shp.TextFrame2.TextRange.Font.Size = 10
    3820 shp.TextFrame2.TextRange.Font.Bold = msoTrue
    3830 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
    3835 shp.TextFrame2.MarginRight = 10
    3840 shp.OnAction = "ToggleHiddenSheets" """
content = content.replace(old_toggle_sheets, "")


# Remove old Exit button and old Toggle Layout
old_exit_toggle = """    ' Add Toggle Layout Button
    Dim shpToggleLayout As Shape
    On Error Resume Next
    wsMain.Shapes("shpToggleLayout").Delete
    On Error GoTo ERR_HANDLER
    Set shpToggleLayout = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("A8").Left + 10, wsMain.Range("A8").Top, 140, 30)
    shpToggleLayout.Name = "shpToggleLayout"
    shpToggleLayout.Fill.ForeColor.RGB = RGB(80, 80, 80)
    shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)
    shpToggleLayout.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpToggleLayout.TextFrame2.TextRange.Font.Size = 11
    shpToggleLayout.TextFrame2.TextRange.Font.Bold = msoTrue
    shpToggleLayout.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shpToggleLayout.OnAction = "ToggleDisplayMode"

    Dim shpExit As Shape
    3880 Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("A19").Left + 60, wsMain.Range("A19").Top, 120, 25)
    3890 shpExit.Fill.ForeColor.RGB = RGB(192, 0, 0)
    3900 shpExit.TextFrame2.TextRange.Text = ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1502) & ChrW(1492) & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514)
    3910 shpExit.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    3920 shpExit.TextFrame2.TextRange.Font.Size = 10
    3930 shpExit.TextFrame2.TextRange.Font.Bold = msoTrue
    3940 shpExit.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    3950 shpExit.Name = "btnNavExit"
    3960 shpExit.OnAction = "ExitSystem" """
content = content.replace(old_exit_toggle, "")


# Add the new system buttons under Main Menu
new_system_buttons = """    ' --- Add System Tools Block under Main Menu ---
    Dim sysW As Double, sysH As Double
    sysW = 300
    sysH = 32
    
    ' 8 - Operation Manual
    Dim shpOpManual As Shape
    On Error Resume Next
    wsMain.Shapes("shpManualMsg").Delete
    On Error GoTo ERR_HANDLER
    Set shpOpManual = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C16").Top, sysW, sysH)
    shpOpManual.Name = "shpManualMsg"
    shpOpManual.Fill.ForeColor.RGB = RGB(0, 100, 200)
    shpOpManual.TextFrame2.TextRange.Text = ChrW(1502) & ChrW(1491) & ChrW(1512) & ChrW(1497) & ChrW(1498) & " " & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
    shpOpManual.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpOpManual.TextFrame2.TextRange.Font.Size = 13
    shpOpManual.TextFrame2.TextRange.Font.Bold = msoTrue
    shpOpManual.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shpOpManual.OnAction = "OpenManualSheet"
    
    ' 9 - Toggle Hidden Sheets
    Dim shpToggle As Shape
    On Error Resume Next
    wsMain.Shapes("shpToggleSheets").Delete
    On Error GoTo ERR_HANDLER
    Set shpToggle = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C18").Top, sysW, sysH)
    shpToggle.Name = "shpToggleSheets"
    shpToggle.Fill.ForeColor.RGB = RGB(80, 80, 80)
    shpToggle.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & "/" & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1512) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514)
    shpToggle.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpToggle.TextFrame2.TextRange.Font.Size = 13
    shpToggle.TextFrame2.TextRange.Font.Bold = msoTrue
    shpToggle.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shpToggle.OnAction = "ToggleHiddenSheets"
    
    ' 10 - Toggle Layout
    Dim shpToggleLayout As Shape
    On Error Resume Next
    wsMain.Shapes("shpToggleLayout").Delete
    On Error GoTo ERR_HANDLER
    Set shpToggleLayout = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C20").Top, sysW, sysH)
    shpToggleLayout.Name = "shpToggleLayout"
    shpToggleLayout.Fill.ForeColor.RGB = RGB(80, 80, 80)
    shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)
    shpToggleLayout.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpToggleLayout.TextFrame2.TextRange.Font.Size = 13
    shpToggleLayout.TextFrame2.TextRange.Font.Bold = msoTrue
    shpToggleLayout.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shpToggleLayout.OnAction = "ToggleDisplayMode"
    
    ' 11 - Exit System
    Dim shpExit As Shape
    On Error Resume Next
    wsMain.Shapes("btnNavExit").Delete
    On Error GoTo ERR_HANDLER
    Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C22").Top, sysW, sysH)
    shpExit.Name = "btnNavExit"
    shpExit.Fill.ForeColor.RGB = RGB(192, 0, 0)
    shpExit.TextFrame2.TextRange.Text = ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1502) & ChrW(1492) & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514)
    shpExit.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpExit.TextFrame2.TextRange.Font.Size = 13
    shpExit.TextFrame2.TextRange.Font.Bold = msoTrue
    shpExit.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shpExit.OnAction = "ExitSystem"
"""

content = content.replace("' ---- Add Status Message Box", new_system_buttons + "\n    ' ---- Add Status Message Box")


# Modify the Phone Number Footer
old_phone = """    ' ---- Add WhatsApp Credit in Row 19 ----
4791 On Error Resume Next
4792 wsMain.Range("A16:L22").UnMerge
4793 With wsMain.Range("C16:I16")"""

new_phone = """    ' ---- Add WhatsApp Credit as Footer ----
4791 On Error Resume Next
4792 wsMain.Range("A16:L26").UnMerge
4793 With wsMain.Range("A24:L24")"""

content = content.replace(old_phone, new_phone)


# Remove the OLD Manual Button creation!
old_manual_btn = """    ' ---- Add MATACH Manual Update Button
    Dim shpManual As Shape
    For Each shpManual In wsMain.Shapes
        If shpManual.Name = "shpManualMsg" Then shpManual.Delete
    Next shpManual
    Set shpManual = wsMain.Shapes.AddShape(msoShapeOval, wsMain.Range("L1").Left + (wsMain.Range("L1").Width - 120) / 2, wsMain.Range("L1").Top + wsMain.Range("L1").Height - 80, 120, 80)
    shpManual.Name = "shpManualMsg"
    shpManual.Fill.ForeColor.RGB = RGB(0, 100, 200)
    With shpManual.TextFrame2.TextRange
        .Text = ChrW(1502) & ChrW(1491) & ChrW(1512) & ChrW(1497) & ChrW(1498) & vbCrLf & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Font.Size = 12
        .Font.Bold = msoTrue
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    shpManual.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shpManual.OnAction = "OpenManualSheet" """

content = content.replace(old_manual_btn, "")


# Fix ToggleDisplayMode: System buttons no longer move!
old_toggle_macro = """        ' 1. Matach Manual Button
        Dim shpManual As Shape
        Set shpManual = wsMain.Shapes("shpMatachManual")
        If Not shpManual Is Nothing Then
            shpManual.Left = wsMain.Range("F15").Left + 20
            shpManual.Top = wsMain.Range("F15").Top + 5
        End If
        
        ' 2. Operation Manual Button (shpManualMsg)
        Dim shpOpManual As Shape
        Set shpOpManual = wsMain.Shapes("shpManualMsg")
        If Not shpOpManual Is Nothing Then
            shpOpManual.Left = wsMain.Range("D12").Left
            shpOpManual.Top = wsMain.Range("D12").Top
        End If
        
        ' 3. Toggle Sheets Button
        Dim shpToggle As Shape
        Set shpToggle = wsMain.Shapes("shpToggleSheets")
        If Not shpToggle Is Nothing Then
            shpToggle.Left = wsMain.Range("A12").Left + 60
            shpToggle.Top = wsMain.Range("A12").Top
        End If
        
        ' 4. Exit Button
        Dim shpExitBtn As Shape
        Set shpExitBtn = wsMain.Shapes("btnNavExit")
        If Not shpExitBtn Is Nothing Then
            shpExitBtn.Left = wsMain.Range("A15").Left + 60
            shpExitBtn.Top = wsMain.Range("A15").Top
        End If
        
        ' 5. Clear old V2.257 texts from bottom
        wsMain.Range("A20:Z30").ClearContents
        
        Dim shpToggleLayout As Shape
        Set shpToggleLayout = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout Is Nothing Then
            shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1512) & ChrW(1495) & ChrW(1489)
        End If"""

new_toggle_macro = """        ' 1. Matach Manual Button
        Dim shpManual As Shape
        Set shpManual = wsMain.Shapes("shpMatachManual")
        If Not shpManual Is Nothing Then
            shpManual.Left = wsMain.Range("G16").Left + (wsMain.Range("G16").Width - shpManual.Width) / 2
            shpManual.Top = wsMain.Range("G16").Top + 5
        End If
        
        Dim shpToggleLayout As Shape
        Set shpToggleLayout = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout Is Nothing Then
            shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1512) & ChrW(1495) & ChrW(1489)
        End If"""

content = content.replace(old_toggle_macro, new_toggle_macro)


old_toggle_macro_desktop = """        ' 1. Matach Manual Button
        Dim shpManual2 As Shape
        Set shpManual2 = wsMain.Shapes("shpMatachManual")
        If Not shpManual2 Is Nothing Then
            shpManual2.Left = wsMain.Range("L1").Left + (wsMain.Range("L1").Width - 120) / 2
            shpManual2.Top = wsMain.Range("L1").Top + wsMain.Range("L1").Height - 80
        End If
        
        ' 2. Operation Manual Button (shpManualMsg)
        Dim shpOpManual2 As Shape
        Set shpOpManual2 = wsMain.Shapes("shpManualMsg")
        If Not shpOpManual2 Is Nothing Then
            shpOpManual2.Left = wsMain.Range("L2").Left + (wsMain.Range("L2").Width - 120) / 2
            shpOpManual2.Top = wsMain.Range("L2").Top + 20
        End If
        
        ' 3. Toggle Sheets Button
        Dim shpToggle2 As Shape
        Set shpToggle2 = wsMain.Shapes("shpToggleSheets")
        If Not shpToggle2 Is Nothing Then
            shpToggle2.Left = wsMain.Range("A14").Left + 60
            shpToggle2.Top = wsMain.Range("A14").Top
        End If
        
        ' 4. Exit Button
        Dim shpExitBtn2 As Shape
        Set shpExitBtn2 = wsMain.Shapes("btnNavExit")
        If Not shpExitBtn2 Is Nothing Then
            shpExitBtn2.Left = wsMain.Range("A19").Left + 60
            shpExitBtn2.Top = wsMain.Range("A19").Top
        End If
        
        ' 5. Clear old V2.257 texts from bottom
        wsMain.Range("A20:Z30").ClearContents
        
        Dim shpToggleLayout2 As Shape
        Set shpToggleLayout2 = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout2 Is Nothing Then
            shpToggleLayout2.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)
        End If"""

new_toggle_macro_desktop = """        ' 1. Matach Manual Button
        Dim shpManual2 As Shape
        Set shpManual2 = wsMain.Shapes("shpMatachManual")
        If Not shpManual2 Is Nothing Then
            shpManual2.Left = wsMain.Range("L1").Left + (wsMain.Range("L1").Width - 120) / 2
            shpManual2.Top = wsMain.Range("L1").Top + wsMain.Range("L1").Height - 80
        End If
        
        Dim shpToggleLayout2 As Shape
        Set shpToggleLayout2 = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout2 Is Nothing Then
            shpToggleLayout2.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)
        End If"""

content = content.replace(old_toggle_macro_desktop, new_toggle_macro_desktop)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.270.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 270')
