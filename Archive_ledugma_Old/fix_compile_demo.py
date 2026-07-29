def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    if 'Public Sub ApplyDemoLockOnOpen()' not in content:
        macro = """
Public Sub ApplyDemoLockOnOpen()
    Dim wsMain As Worksheet
    On Error Resume Next
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If wsMain Is Nothing Then Exit Sub
    
    Dim isDemoMode As Boolean
    Dim demoParam As String
    isDemoMode = FORCE_DEMO_MODE
    demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
    If Not FORCE_DEMO_MODE Then
        If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
        If demoParam = ChrW(1500) & ChrW(1488) Or demoParam = "NO" Then isDemoMode = False
    End If
    
    If isDemoMode Then
        wsMain.Range("G3").Value = 2024
        wsMain.Range("G4").Value = 2025
        With wsMain.Range("G3:G4").Validation
            .Delete
        End With
        wsMain.Range("G3:G4").Interior.ColorIndex = 15 ' Grey
        
        ' Add transparent shape over G3:G4 to block clicking
        On Error Resume Next
        wsMain.Shapes("shpDemoLockG3G4").Delete
        On Error GoTo 0
        Dim shpLock As Shape
        Set shpLock = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("G3:G4").Left, wsMain.Range("G3:G4").Top, wsMain.Range("G3:G4").Width, wsMain.Range("G3:G4").Height)
        shpLock.Name = "shpDemoLockG3G4"
        shpLock.Fill.Transparency = 1#
        shpLock.Line.Visible = msoFalse
        shpLock.OnAction = "DemoModeRestricted"
    Else
        ' Restore validation and remove lock shape if not demo
        With wsMain.Range("G3:G4").Validation
            .Delete
            On Error Resume Next
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_years"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
        On Error Resume Next
        wsMain.Shapes("shpDemoLockG3G4").Delete
        wsMain.Shapes("shpDemoMsgText").Delete
        wsMain.Range("D19:K20").ClearContents
        On Error GoTo 0
    End If
End Sub
"""
        with open(bas_file, 'a', encoding='utf-8') as f:
            f.write("\n" + macro)
        print("Fixed Compile Error ApplyDemoLockOnOpen!")
    else:
        print("Macro already exists.")

if __name__ == "__main__":
    main()
