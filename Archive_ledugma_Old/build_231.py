import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.230.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.230', 'VERSION: V2.231')
content = content.replace('Error in V2.230!', 'Error in V2.231!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_230"', 'Attribute VB_Name = "Goren_Claude_V2_231"')
content = content.replace('APP_VERSION As String = "2.230"', 'APP_VERSION As String = "2.231"')
content = content.replace('APP_VERSION = "2.230"', 'APP_VERSION = "2.231"')

start_str = 'Public Sub ApplyDemoLockOnOpen()'
end_str = 'Public Sub OpenManualSheet()'

start_idx = content.find(start_str)
end_idx = content.find(end_str)

new_sub = """Public Sub ApplyDemoLockOnOpen()
    Dim wsMain As Worksheet
    On Error Resume Next
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If wsMain Is Nothing Then Exit Sub

    ' Delete Hyperlink from shpInstallMsg so Desktop uses OnAction
    Dim hl As Hyperlink
    For Each hl In wsMain.Hyperlinks
        If hl.Shape.Name = "shpInstallMsg" Then hl.Delete
    Next hl

    ' Determine Demo Mode
    Dim isDemoMode As Boolean
    Dim demoParam As String
    isDemoMode = FORCE_DEMO_MODE
    demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
    If Not FORCE_DEMO_MODE Then
        If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
        If demoParam = ChrW(1500) & ChrW(1488) Or demoParam = "NO" Then isDemoMode = False
    End If

    ' Delete previous UI elements so we can reconstruct them cleanly
    On Error Resume Next
    wsMain.Shapes("shpDemoMsgText").Delete
    wsMain.Shapes("shpInstallMsg").Delete
    wsMain.Shapes("shpManualMsg").Delete
    wsMain.Shapes("shpDemoLockG3G4").Delete
    On Error GoTo 0
    
    ' Fix white cells if they were broken by previous versions, WITHOUT clearing contents
    wsMain.Range("B13:K15").Interior.Color = RGB(220, 240, 220)
    wsMain.Range("D18:K20").Interior.Color = RGB(220, 240, 220)
    
    ' Always Clear D19:K20 contents to remove old demo text, but KEEP E18 intact!
    wsMain.Range("D19:K20").ClearContents

    ' --- Add Left Button: Operating Manual (Oval shape opening sheet) ---
    Dim shpManual As Shape
    Set shpManual = wsMain.Shapes.AddShape(msoShapeOval, wsMain.Range("L2").Left, wsMain.Range("L2").Top, 120, 80)
    shpManual.Name = "shpManualMsg"
    shpManual.Fill.ForeColor.RGB = RGB(0, 100, 200)
    With shpManual.TextFrame2.TextRange
        .Text = ChrW(1502) & ChrW(1491) & ChrW(1512) & ChrW(1497) & ChrW(1498) & vbCrLf & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
        .Font.Size = 12
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    shpManual.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shpManual.OnAction = "OpenManualSheet"

    ' --- Branch on Demo vs Full Mode ---
    If isDemoMode Then
        wsMain.Range("G3").Value = 2024
        wsMain.Range("G4").Value = 2025
        On Error Resume Next
        wsMain.Range("G3:G4").Validation.Delete
        On Error GoTo 0
        wsMain.Range("G3:G4").Interior.ColorIndex = 15 ' Grey
        
        Dim shpLock As Shape
        Set shpLock = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("G3:G4").Left, wsMain.Range("G3:G4").Top, wsMain.Range("G3:G4").Width, wsMain.Range("G3:G4").Height)
        shpLock.Name = "shpDemoLockG3G4"
        shpLock.Fill.Transparency = 1#
        shpLock.Line.Visible = msoFalse
        shpLock.OnAction = "DemoModeRestricted"
        
        wsMain.Range("A1").Font.Color = RGB(200, 0, 0)
        
        ' --- Add Right Button: Installation Instructions (SharePoint) ONLY IN DEMO MODE ---
        Dim shpInstall As Shape
        Set shpInstall = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("A2").Left + 10, wsMain.Range("A2").Top, 220, 65)
        shpInstall.Name = "shpInstallMsg"
        shpInstall.Fill.ForeColor.RGB = RGB(245, 245, 245)
        shpInstall.Line.ForeColor.RGB = RGB(0, 0, 139)
        shpInstall.Line.Weight = 2
        With shpInstall.TextFrame2.TextRange
            .Text = ChrW(1492) & ChrW(1511) & ChrW(1513) & ChrW(32) & ChrW(1499) & ChrW(1488) & ChrW(1503) & ChrW(32) & ChrW(1500) & ChrW(1508) & ChrW(1514) & ChrW(1497) & ChrW(1495) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & vbCrLf & ChrW(1504) & ChrW(1497) & ChrW(1514) & ChrW(1503) & ChrW(32) & ChrW(1500) & ChrW(1513) & ChrW(1500) & ChrW(1493) & ChrW(1495) & ChrW(32) & "WhatsApp" & ChrW(32) & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & ChrW(32) & "054-6677396" & ChrW(32) & ChrW(1500) & ChrW(1511) & ChrW(1489) & ChrW(1500) & ChrW(1514) & ChrW(32) & ChrW(1506) & ChrW(1494) & ChrW(1512) & ChrW(1492)
            .Font.Size = 10
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 139)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        shpInstall.TextFrame2.VerticalAnchor = msoAnchorMiddle
        wsMain.Hyperlinks.Add Anchor:=shpInstall, Address:="https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQBnB0klIujxR4O5fALynO8EAWddhgVppMJI7THxhW3R6fo?e=Z8iVEV"
        
    Else
        ' FULL MODE
        If IsEmpty(wsMain.Range("G3").Value) Then wsMain.Range("G3").Value = 2024
        If IsEmpty(wsMain.Range("G4").Value) Then wsMain.Range("G4").Value = 2025
        wsMain.Range("G3:G4").Interior.Color = RGB(255, 245, 230)
        On Error Resume Next
        With wsMain.Range("G3:G4").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_years"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        On Error GoTo 0
        
        wsMain.Range("A1").Font.Color = RGB(0, 0, 0)
    End If
    
    wsMain.Range("A22").Value = ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " v" & APP_VERSION
    wsMain.Range("A22").Font.Size = 10
    wsMain.Range("A22").Font.Color = RGB(150, 150, 150)
    
    On Error Resume Next
    Application.Goto wsMain.Range("F10")
    wsMain.Unprotect "Z961814r"
    wsMain.Protect Password:="Z961814r", UserInterfaceOnly:=True
    
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
    On Error GoTo 0
    
    Application.ScreenUpdating = True
End Sub

"""

content = content[:start_idx] + new_sub + content[end_idx:]

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.231.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 231')
