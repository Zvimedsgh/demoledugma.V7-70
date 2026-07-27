import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.116.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_if_demo = """If isDemoMode Then
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
1775 On Error Resume Next
.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_years"
.IgnoreBlank = True
.InCellDropdown = True
End With
wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
On Error Resume Next
wsMain.Shapes("shpDemoLockG3G4").Delete
On Error GoTo 0
End If"""

new_if_demo = """If isDemoMode Then
wsMain.Range("G3").Value = 2024
wsMain.Range("G4").Value = 2025
With wsMain.Range("G3:G4").Validation
.Delete
End With
wsMain.Range("G3:G4").Interior.ColorIndex = 15 ' Grey

' Add transparent shape over G3:G4 to block clicking
On Error Resume Next
wsMain.Shapes("shpDemoLockG3G4").Delete
wsMain.Shapes("shpDemoContact").Delete
On Error GoTo 0

Dim shpLock As Shape
Set shpLock = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("G3:G4").Left, wsMain.Range("G3:G4").Top, wsMain.Range("G3:G4").Width, wsMain.Range("G3:G4").Height)
shpLock.Name = "shpDemoLockG3G4"
shpLock.Fill.Transparency = 1#
shpLock.Line.Visible = msoFalse
shpLock.OnAction = "DemoModeRestricted"

' Add Demo contact message
Dim shpDemoContact As Shape
Set shpDemoContact = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("D15").Left + 50, wsMain.Range("D15").Top + 5, 400, 30)
shpDemoContact.Name = "shpDemoContact"
shpDemoContact.Fill.Visible = msoFalse
shpDemoContact.Line.Visible = msoFalse
shpDemoContact.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " 054-6677396"
shpDemoContact.TextFrame2.TextRange.Font.Size = 16
shpDemoContact.TextFrame2.TextRange.Font.Bold = msoTrue
shpDemoContact.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 112, 192)
shpDemoContact.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

Else
' Restore validation and remove lock shape if not demo
With wsMain.Range("G3:G4").Validation
.Delete
1775 On Error Resume Next
.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_years"
.IgnoreBlank = True
.InCellDropdown = True
End With
wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
On Error Resume Next
wsMain.Shapes("shpDemoLockG3G4").Delete
wsMain.Shapes("shpDemoContact").Delete
On Error GoTo 0
End If"""

content = content.replace(old_if_demo, new_if_demo)


# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_116"', 'Attribute VB_Name = "Goren_Claude_V2_117"')
content = content.replace('VERSION: V2.116', 'VERSION: V2.117')
content = content.replace('APP_VERSION As String = "2.116"', 'APP_VERSION As String = "2.117"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.117.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.117 created.")
