import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.265.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.265"', 'APP_VERSION As String = "2.266"')
content = content.replace('VERSION: V2.265', 'VERSION: V2.266')
content = content.replace('Error in V2.265!', 'Error in V2.266!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_265"', 'Attribute VB_Name = "Goren_Claude_V2_266"')

# Changelog
changelog = """' CHANGES IN 2.266:
'   - UI: Added Dynamic Layout Toggle (Option 4) for Laptop Responsive Design!
"""
content = content.replace("' CHANGES IN 2.265:", changelog + "' CHANGES IN 2.265:")

# Add the new button creation in A00_SetupMainSheet
# Let's find where shpExit is added and add it right before.
old_exit_btn = '''    Dim shpExit As Shape
    3880 Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("A19").Left + 60, wsMain.Range("A19").Top, 120, 25)'''

new_toggle_btn = '''    ' Add Toggle Layout Button
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
    3880 Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("A19").Left + 60, wsMain.Range("A19").Top, 120, 25)'''

content = content.replace(old_exit_btn, new_toggle_btn)


toggle_macro = """
' -------------------------------------------------------------------------
' MACRO: ToggleDisplayMode (Dynamic Layout for Laptops)
' -------------------------------------------------------------------------
Public Sub ToggleDisplayMode()
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(ChrW(1491) & ChrW(1508) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514))
    
    If wsMain Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    wsMain.Unprotect Password:="1234"
    
    Dim isDesktop As Boolean
    If wsMain.Range("J2").Value = ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506) Then
        isDesktop = True
    Else
        isDesktop = False
    End If
    
    If isDesktop Then
        ' Move from Desktop to Laptop
        wsMain.Range("J2:K4").ClearContents
        wsMain.Range("J2:K4").Interior.Color = xlNone
        wsMain.Range("J2:K4").Borders.LineStyle = xlNone
        
        wsMain.Range("F12").Value = ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506)
        wsMain.Range("G12").Value = ChrW(1513) & ChrW(1506) & ChrW(1512)
        wsMain.Range("F12:G12").Interior.Color = RGB(0, 100, 0)
        wsMain.Range("F12:G12").Font.Color = RGB(255, 255, 255)
        wsMain.Range("F12:G12").Font.Bold = True
        wsMain.Range("F12:G12").Font.Size = 12
        
        wsMain.Range("F13").Value = ChrW(1491) & ChrW(1493) & ChrW(1500) & ChrW(1512)
        wsMain.Range("F14").Value = ChrW(1488) & ChrW(1497) & ChrW(1512) & ChrW(1493)
        wsMain.Range("F13:F14").Interior.Color = RGB(255, 255, 204)
        wsMain.Range("F13:F14").Font.Bold = True
        wsMain.Range("F13:F14").Font.Size = 11
        
        wsMain.Range("G13").Interior.Color = RGB(230, 240, 255)
        wsMain.Range("G14").Interior.Color = RGB(230, 240, 255)
        wsMain.Range("G13:G14").Font.Bold = True
        wsMain.Range("G13:G14").Font.Size = 11
        wsMain.Range("G13:G14").NumberFormat = "0.0000"
        
        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("G13")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("G14")
        
        Dim shpManual As Shape
        Set shpManual = wsMain.Shapes("shpMatachManual")
        If Not shpManual Is Nothing Then
            shpManual.Left = wsMain.Range("A13").Left + 60
            shpManual.Top = wsMain.Range("A13").Top
        End If
        
        Dim shpToggleLayout As Shape
        Set shpToggleLayout = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout Is Nothing Then
            shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1512) & ChrW(1495) & ChrW(1489)
        End If
        
        Call get_rates
        
    Else
        ' Move from Laptop to Desktop
        wsMain.Range("F12:G14").ClearContents
        wsMain.Range("F12:G14").Interior.Color = xlNone
        wsMain.Range("F12:G14").Borders.LineStyle = xlNone
        
        wsMain.Range("J2").Value = ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506)
        wsMain.Range("K2").Value = ChrW(1513) & ChrW(1506) & ChrW(1512)
        wsMain.Range("J2:K2").Interior.Color = RGB(0, 100, 0)
        wsMain.Range("J2:K2").Font.Color = RGB(255, 255, 255)
        wsMain.Range("J2:K2").Font.Bold = True
        wsMain.Range("J2:K2").Font.Size = 12
        
        wsMain.Range("J3").Value = ChrW(1491) & ChrW(1493) & ChrW(1500) & ChrW(1512)
        wsMain.Range("J4").Value = ChrW(1488) & ChrW(1497) & ChrW(1512) & ChrW(1493)
        wsMain.Range("J3:J4").Interior.Color = RGB(255, 255, 204)
        wsMain.Range("J3:J4").Font.Bold = True
        wsMain.Range("J3:J4").Font.Size = 11
        
        wsMain.Range("K3").Interior.Color = RGB(230, 240, 255)
        wsMain.Range("K4").Interior.Color = RGB(230, 240, 255)
        wsMain.Range("K3:K4").Font.Bold = True
        wsMain.Range("K3:K4").Font.Size = 11
        wsMain.Range("K3:K4").NumberFormat = "0.0000"
        
        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("K3")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("K4")
        
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
        End If
        
        Call get_rates
    End If
    
    wsMain.Protect Password:="1234", UserInterfaceOnly:=True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
"""

content = content + "\n\n" + toggle_macro

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.266.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 266')
