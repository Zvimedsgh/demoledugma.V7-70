import re

def main():
    with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.225.bas', 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Update version
    content = content.replace('Error in V2.225!', 'Error in V2.226!')
    content = content.replace('Attribute VB_Name = "Goren_Claude_V2_225"', 'Attribute VB_Name = "Goren_Claude_V2_226"')
    content = content.replace('VERSION: V2.225', 'VERSION: V2.226')

    # 2. Update default years in A00_SetupMainSheet
    content = content.replace('1550 If IsEmpty(wsMain.Range("G3").Value) Then wsMain.Range("G3").Value = 2025', '1550 wsMain.Range("G3").Value = 2024')
    content = content.replace('1560 If IsEmpty(wsMain.Range("G4").Value) Then wsMain.Range("G4").Value = 2024', '1560 wsMain.Range("G4").Value = 2025')

    # 3. Completely replace ApplyDemoLockOnOpen
    start_str = 'Public Sub ApplyDemoLockOnOpen()'
    end_str = 'Public Sub EmailNewClients()'

    start_idx = content.find(start_str)
    end_idx = content.find(end_str)

    if start_idx != -1 and end_idx != -1:
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
    wsMain.Shapes("shpDemoLockG3G4").Delete
    wsMain.Range("B14:K15").Clear
    wsMain.Range("D19:K20").Clear
    On Error GoTo 0

    ' --- Add Operating Manual Button (For both Demo and Full) ---
    Dim shpInstall As Shape
    Set shpInstall = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("A2").Left + 10, wsMain.Range("A2").Top, 220, 65)
    shpInstall.Name = "shpInstallMsg"
    shpInstall.Fill.ForeColor.RGB = RGB(245, 245, 245)
    shpInstall.Line.ForeColor.RGB = RGB(0, 0, 139)
    shpInstall.Line.Weight = 2
    With shpInstall.TextFrame2.TextRange
        .Text = ChrW(1500) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(32) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & ChrW(32) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(32) & ChrW(1511) & ChrW(1500) & ChrW(1511) & ChrW(32) & ChrW(1499) & ChrW(1488) & ChrW(1503) & ChrW(32) & ChrW(1511) & ChrW(1500) & ChrW(1497) & ChrW(1511) & ChrW(32) & ChrW(1497) & ChrW(1502) & ChrW(1504) & ChrW(1497) & ChrW(32) & ChrW(1508) & vbCrLf & "054-6677396" & vbCrLf & ChrW(1513) & ChrW(1500) & ChrW(1495) & ChrW(32) & ChrW(1493) & ChrW(1493) & ChrW(1488) & ChrW(1496) & ChrW(1505) & ChrW(1488) & ChrW(1508) & ChrW(32) & ChrW(1500) & ChrW(1506) & ChrW(1494) & ChrW(1512) & ChrW(1492)
        .Font.Size = 12
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(0, 0, 139)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    shpInstall.TextFrame2.VerticalAnchor = msoAnchorMiddle

    ' --- Branch on Demo vs Full Mode ---
    If isDemoMode Then
        ' Demo Mode Restrictions
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
        
        wsMain.Range("A1").Font.Color = RGB(200, 0, 0) ' Red for Demo
    Else
        ' Full Mode Restorations
        wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
        ' Restore Validation
        On Error Resume Next
        With wsMain.Range("G3:G4").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_years"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        On Error GoTo 0
        
        wsMain.Range("A1").Font.Color = RGB(0, 0, 0) ' Black for Real
    End If
    
    ' Navigation and UI Protection
    On Error Resume Next
    Application.Goto wsMain.Range("F10")
    ' DO NOT Protect sheet here unless absolutely necessary, or ensure we only protect UI.
    ' If we protect it here, we must unlock G3:G10 first! Actually, A00_SetupMainSheet handles all that.
    ' We will just unprotect and then protect with UserInterfaceOnly to be safe.
    wsMain.Unprotect "Z961814r"
    wsMain.Protect Password:="Z961814r", UserInterfaceOnly:=True
    
    ' Navigate to home page
    ThisWorkbook.Worksheets(homeSheetName).Activate
    
    ' But if operation instruction sheet is enabled, jump to it instead
    Dim sOp2 As String
    sOp2 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
    If UCase$(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("AA2").Value) <> "YES" Then
        ThisWorkbook.Worksheets(sOp2).Activate
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True
End Sub

' ============================================================================
' EMAIL NEW CLIENTS: Export visible rows to temp Excel file and send via Outlook
"""
        
        content = content[:start_idx] + new_sub + content[end_idx:]

    with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.226.bas', 'w', encoding='utf-8') as f:
        f.write(content)
    print('Generated 226')

if __name__ == '__main__':
    main()
