import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.13.bas', 'r', encoding='utf-8') as f:
    content = f.read()

new_demo_lock = '''Public Sub ApplyDemoLockOnOpen()
    Dim wsMain As Worksheet
    On Error Resume Next
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If wsMain Is Nothing Then Exit Sub

    ' Delete Hyperlink from shpInstallMsg so Desktop uses OnAction
    Dim hl As Hyperlink
    For Each hl In wsMain.Hyperlinks
        ' If hl.Shape.Name = "shpInstallMsg" Then hl.Delete
    Next hl

    ' Determine Demo Mode using helper function
    Dim isDemoMode As Boolean
    isDemoMode = IsSystemInDemoMode()

    ' Delete previous UI elements so we can reconstruct them cleanly
    On Error Resume Next
    wsMain.Shapes("shpDemoMsgText").Delete
    wsMain.Shapes("shpDemoLockG3G4").Delete
    On Error GoTo 0

    ' Fix white cells if they were broken by previous versions, WITHOUT clearing contents
    wsMain.Range("B13:K15").Interior.Color = RGB(220, 240, 220)
    wsMain.Range("D18:K20").Interior.Color = RGB(220, 240, 220)

    ' Always Clear D19:K20 contents to remove old demo text, but KEEP E18 intact!
    wsMain.Range("J19:K20").ClearContents

    ' --- Branch on Demo vs Full Mode ---
    If isDemoMode Then
        wsMain.Range("G3").Value = 2024
        wsMain.Range("G4").Value = 2025
        On Error Resume Next
        wsMain.Range("G3").MergeArea.Validation.Delete
        wsMain.Range("G4").MergeArea.Validation.Delete
        On Error GoTo 0
        wsMain.Range("G3").MergeArea.Interior.ColorIndex = 15
        wsMain.Range("G4").MergeArea.Interior.ColorIndex = 15

        Dim shpLock As Shape
        Set shpLock = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("G3").Left, wsMain.Range("G3").Top, wsMain.Range("G3").Width, wsMain.Range("G3").Height + wsMain.Range("G4").Height)
        shpLock.Name = "shpDemoLockG3G4"
        shpLock.Fill.Transparency = 1#
        shpLock.Line.Visible = msoFalse
        shpLock.OnAction = "DemoModeRestricted"

        wsMain.Range("A1").Font.Color = RGB(200, 0, 0)

    Else
        ' FULL MODE
        If IsEmpty(wsMain.Range("G3").Value) Then wsMain.Range("G3").Value = 2024
        If IsEmpty(wsMain.Range("G4").Value) Then wsMain.Range("G4").Value = 2025
        wsMain.Range("G3").MergeArea.Interior.Color = RGB(255, 245, 230)
        wsMain.Range("G4").MergeArea.Interior.Color = RGB(255, 245, 230)
        On Error Resume Next
        With wsMain.Range("G3").MergeArea.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="2020,2021,2022,2023,2024,2025"
            .InCellDropdown = True
        End With
        With wsMain.Range("G4").MergeArea.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="2020,2021,2022,2023,2024,2025"
            .InCellDropdown = True
        End With
        On Error GoTo 0
        wsMain.Range("A1").Font.Color = RGB(0, 0, 0)
    End If
End Sub'''

pattern_demo_lock = re.compile(r'Public Sub ApplyDemoLockOnOpen\(\).*?End Sub', re.DOTALL)
content = pattern_demo_lock.sub(new_demo_lock, content, count=1)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.13.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Done step 3')
