import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_str = """        With wsMain.Range("G3:G4").Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=FALSE"
            .ErrorMessage = ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497) & " " & ChrW(1488) & ChrW(1508) & ChrW(1513) & ChrW(1512) & ChrW(1497) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492)
            .ShowError = True
        End With
        wsMain.Range("G3:G4").Interior.ColorIndex = 15 ' Grey
    Else
        ' Remove validation if not demo
        wsMain.Range("G3:G4").Validation.Delete
        wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
    End If
End Sub"""

new_str = """        With wsMain.Range("G3:G4").Validation
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
        ' Remove validation and lock shape if not demo
        wsMain.Range("G3:G4").Validation.Delete
        wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
        On Error Resume Next
        wsMain.Shapes("shpDemoLockG3G4").Delete
        On Error GoTo 0
    End If
End Sub"""

if old_str in content:
    content = content.replace(old_str, new_str)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Fixed ApplyDemoLockOnOpen")
else:
    print("Could not find ApplyDemoLockOnOpen string")
