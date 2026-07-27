import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.161.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_macro = """
Public Sub FixUIButtons()
    ' One-time macro to fix and reorganize the UI buttons on the home sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    ws.Unprotect "Z961814r"
    
    Dim shp As Shape
    Dim btn1 As Shape, btn2 As Shape, btn3 As Shape, btn4 As Shape
    Dim btn5 As Shape, btn6 As Shape, btn7 As Shape
    
    ' Find existing buttons by their TEXT (much more reliable than OnAction)
    For Each shp In ws.Shapes
        If TypeName(shp) <> "GroupObject" And shp.Type <> msoComment Then
            On Error Resume Next
            Dim txt As String
            txt = ""
            txt = shp.TextFrame.Characters.Text
            On Error GoTo 0
            
            If Left(txt, 4) = "1 - " Then Set btn1 = shp
            If Left(txt, 4) = "2 - " Then Set btn2 = shp
            If Left(txt, 4) = "3 - " Then
                If InStr(txt, ChrW(1491) & ChrW(1497) & ChrW(1493) & ChrW(1493) & ChrW(1495)) > 0 Then
                    ' "דיווח" - This is the NEW button 3!
                    Set btn3 = shp
                Else
                    ' Old button 3 (מצגת) -> this should be button 4
                    Set btn4 = shp
                End If
            End If
            If Left(txt, 4) = "4 - " Then Set btn5 = shp
            If Left(txt, 4) = "5 - " Then Set btn6 = shp
            If Left(txt, 4) = "6 - " Then Set btn7 = shp
        End If
    Next shp
    
    ' If button 3 doesn't exist yet, duplicate button 2
    If btn3 Is Nothing And Not btn2 Is Nothing Then
        Set btn3 = btn2.Duplicate
        ' We must adjust its OnAction so it calls the right thing!
        btn3.OnAction = "ShowResultSheets"
    End If
    
    ' Update OnActions just in case they were lost!
    If Not btn1 Is Nothing Then btn1.OnAction = "BuildReview"
    If Not btn2 Is Nothing Then btn2.OnAction = "ApplyCorrectionsAndBuildReports"
    If Not btn3 Is Nothing Then btn3.OnAction = "ShowResultSheets"
    If Not btn4 Is Nothing Then btn4.OnAction = "BuildPresentation"
    If Not btn5 Is Nothing Then btn5.OnAction = "SaveReportsToFolder"
    If Not btn6 Is Nothing Then btn6.OnAction = "ViewReportsFolder"
    If Not btn7 Is Nothing Then btn7.OnAction = "NewClients"
    
    ' Update text for all buttons
    If Not btn1 Is Nothing Then btn1.TextFrame.Characters.Text = "1 - " & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1514) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
    If Not btn2 Is Nothing Then btn2.TextFrame.Characters.Text = "2 - " & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(34) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    If Not btn3 Is Nothing Then btn3.TextFrame.Characters.Text = "3 - " & ChrW(1492) & ChrW(1510) & ChrW(1490) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1491) & ChrW(1497) & ChrW(1493) & ChrW(1493) & ChrW(1495)
    If Not btn4 Is Nothing Then btn4.TextFrame.Characters.Text = "4 - " & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1493) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
    If Not btn5 Is Nothing Then btn5.TextFrame.Characters.Text = "5 - " & ChrW(1513) & ChrW(1502) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    If Not btn6 Is Nothing Then btn6.TextFrame.Characters.Text = "6 - " & ChrW(1510) & ChrW(1508) & ChrW(1497) & ChrW(1497) & ChrW(1492) & " " & ChrW(1489) & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1513) & ChrW(1502) & ChrW(1493) & ChrW(1512) & ChrW(1497) & ChrW(1501)
    If Not btn7 Is Nothing Then btn7.TextFrame.Characters.Text = "7 - " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
    
    ' Position them nicely (vertical stack)
    Dim startTop As Single
    Dim spacing As Single
    Dim bLeft As Single
    Dim bHeight As Single
    
    If Not btn1 Is Nothing And Not btn2 Is Nothing Then
        startTop = btn1.Top
        spacing = btn2.Top - (btn1.Top + btn1.Height)
        bLeft = btn1.Left
        bHeight = btn1.Height
    Else
        startTop = 40
        spacing = 15
        bLeft = 600
        bHeight = 30
    End If
    
    If spacing < 5 Then spacing = 10
    
    Dim curTop As Single
    curTop = startTop
    
    If Not btn1 Is Nothing Then
        btn1.Top = curTop
        btn1.Left = bLeft
        curTop = curTop + bHeight + spacing
    End If
    If Not btn2 Is Nothing Then
        btn2.Top = curTop
        btn2.Left = bLeft
        curTop = curTop + bHeight + spacing
    End If
    If Not btn3 Is Nothing Then
        btn3.Top = curTop
        btn3.Left = bLeft
        curTop = curTop + bHeight + spacing
    End If
    If Not btn4 Is Nothing Then
        btn4.Top = curTop
        btn4.Left = bLeft
        curTop = curTop + bHeight + spacing
    End If
    If Not btn5 Is Nothing Then
        btn5.Top = curTop
        btn5.Left = bLeft
        curTop = curTop + bHeight + spacing
    End If
    If Not btn6 Is Nothing Then
        btn6.Top = curTop
        btn6.Left = bLeft
        curTop = curTop + bHeight + spacing
    End If
    If Not btn7 Is Nothing Then
        btn7.Top = curTop
        btn7.Left = bLeft
    End If
    
    ws.Protect UserInterfaceOnly:=True
    MsgBox "Buttons updated successfully!", vbInformation
End Sub
"""

# Replace old FixUIButtons block
start_idx = -1
end_idx = -1
for i, line in enumerate(lines):
    if "Public Sub FixUIButtons()" in line:
        start_idx = i
    if "MsgBox \"Buttons updated successfully!\", vbInformation" in line and start_idx != -1:
        end_idx = i + 2
        break

if start_idx != -1 and end_idx != -1:
    new_lines = lines[:start_idx]
    new_lines.append(new_macro)
    new_lines.extend(lines[end_idx:])
    
    # Update version to 2.162
    for i in range(len(new_lines)):
        new_lines[i] = new_lines[i].replace('Attribute VB_Name = "Goren_Claude_V2_161"', 'Attribute VB_Name = "Goren_Claude_V2_162"')
        new_lines[i] = new_lines[i].replace('VERSION: V2.161', 'VERSION: V2.162')
        new_lines[i] = new_lines[i].replace('APP_VERSION As String = "2.161"', 'APP_VERSION As String = "2.162"')
        
    with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.162.bas', 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    print("Created V2.162 with reliable FixUIButtons.")
else:
    print("Could not find FixUIButtons block.")

