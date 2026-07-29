import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.226.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.226', 'VERSION: V2.227')
content = content.replace('Error in V2.227!', 'Error in V2.227!')
content = content.replace('Error in V2.226!', 'Error in V2.227!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_226"', 'Attribute VB_Name = "Goren_Claude_V2_227"')

start_str = 'Public Sub ApplyDemoLockOnOpen()'
end_str = 'Public Sub AssignButtonMacros()'

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
    wsMain.Range("B14:K15").Clear
    wsMain.Range("D19:K20").Clear
    On Error GoTo 0

    ' --- Add Right Button: Installation Instructions (SharePoint) ---
    Dim shpInstall As Shape
    Set shpInstall = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("A2").Left + 10, wsMain.Range("A2").Top, 220, 65)
    shpInstall.Name = "shpInstallMsg"
    shpInstall.Fill.ForeColor.RGB = RGB(245, 245, 245)
    shpInstall.Line.ForeColor.RGB = RGB(0, 0, 139)
    shpInstall.Line.Weight = 2
    With shpInstall.TextFrame2.TextRange
        ' "Hakesh kan liftechat horot ha'fa'la..."
        .Text = ChrW(1492) & ChrW(1511) & ChrW(1513) & ChrW(32) & ChrW(1499) & ChrW(1488) & ChrW(1503) & ChrW(32) & ChrW(1500) & ChrW(1508) & ChrW(1514) & ChrW(1497) & ChrW(1495) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & vbCrLf & ChrW(1504) & ChrW(1497) & ChrW(1514) & ChrW(1503) & ChrW(32) & ChrW(1500) & ChrW(1513) & ChrW(1500) & ChrW(1493) & ChrW(1495) & ChrW(32) & "WhatsApp" & ChrW(32) & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & ChrW(32) & "054-6677396" & ChrW(32) & ChrW(1500) & ChrW(1511) & ChrW(1489) & ChrW(1500) & ChrW(1514) & ChrW(32) & ChrW(1506) & ChrW(1494) & ChrW(1512) & ChrW(1492) & ChrW(32) & ChrW(1489) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
        .Font.Size = 10
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(0, 0, 139)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    shpInstall.TextFrame2.VerticalAnchor = msoAnchorMiddle
    wsMain.Hyperlinks.Add Anchor:=shpInstall, Address:="https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQBnB0klIujxR4O5fALynO8EAWddhgVppMJI7THxhW3R6fo?e=Z8iVEV"

    ' --- Add Left Button: Operating Manual (Diamond shape opening sheet) ---
    Dim shpManual As Shape
    ' Top Left, let's put it at M3 for example
    Set shpManual = wsMain.Shapes.AddShape(msoShapeDiamond, wsMain.Range("L2").Left, wsMain.Range("L2").Top, 120, 80)
    shpManual.Name = "shpManualMsg"
    shpManual.Fill.ForeColor.RGB = RGB(220, 100, 0)
    With shpManual.TextFrame2.TextRange
        ' "Madrich Tifool"
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
    Else
        If IsEmpty(wsMain.Range("G3").Value) Then wsMain.Range("G3").Value = 2024
        If IsEmpty(wsMain.Range("G4").Value) Then wsMain.Range("G4").Value = 2025
        wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
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
    
    On Error Resume Next
    Application.Goto wsMain.Range("F10")
    wsMain.Unprotect "Z961814r"
    wsMain.Protect Password:="Z961814r", UserInterfaceOnly:=True
    
    ThisWorkbook.Worksheets(homeSheetName).Activate
    On Error GoTo 0
    
    Application.ScreenUpdating = True
End Sub

Public Sub OpenManualSheet()
    On Error Resume Next
    Dim sOp As String
    sOp = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
    ThisWorkbook.Worksheets(sOp).Activate
    On Error GoTo 0
End Sub

"""

content = content[:start_idx] + new_sub + content[end_idx:]

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.227.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 227')
