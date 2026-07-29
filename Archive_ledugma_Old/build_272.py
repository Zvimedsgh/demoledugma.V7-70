import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.271.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False

for line in lines:
    if 'Public Sub ToggleDisplayMode()' in line:
        skip = True
    
    if skip and 'End Sub' in line:
        skip = False
        continue
        
    if not skip:
        new_lines.append(line)

content = "".join(new_lines)

# Fix the manual button deletion
content = content.replace('wsMain.Shapes("shpManualMsg").Delete', '\' wsMain.Shapes("shpManualMsg").Delete \' KEPT FOR V2.272')

# Rewrite ToggleDisplayMode using .Cut
new_toggle_macro = """Public Sub ToggleDisplayMode()
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(ChrW(1491) & ChrW(1508) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514))
    
    If wsMain Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    wsMain.Unprotect Password:="1234"
    
    Dim isDesktop As Boolean
    ' Check if Matach table is at Desktop position (J2)
    If wsMain.Range("J2").Value = ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506) Then
        isDesktop = True
    Else
        isDesktop = False
    End If
    
    If isDesktop Then
        ' Move from Desktop to Laptop
        ' Cut Matach table (J2:K7) and move it to F14:G19
        wsMain.Range("J2:K7").Cut Destination:=wsMain.Range("F14")
        
        ' Update Named Ranges
        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("G15")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("G16")
        
        ' 1. Matach Manual Button
        Dim shpManual As Shape
        Set shpManual = wsMain.Shapes("shpMatachManual")
        If Not shpManual Is Nothing Then
            shpManual.Left = wsMain.Range("G18").Left + (wsMain.Range("G18").Width - shpManual.Width) / 2
            shpManual.Top = wsMain.Range("G18").Top + 5
        End If
        
        wsMain.Range("A25:Z30").ClearContents
        
        Dim shpToggleLayout As Shape
        Set shpToggleLayout = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout Is Nothing Then
            shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1512) & ChrW(1495) & ChrW(1489)
        End If
        
    Else
        ' Move from Laptop to Desktop
        ' Cut Matach table (F14:G19) and move it back to J2:K7
        wsMain.Range("F14:G19").Cut Destination:=wsMain.Range("J2")
        
        ' Restore background color for demo message if needed
        If wsMain.Range("B13").Interior.Color = RGB(220, 240, 220) Then
            wsMain.Range("B13:K15").Interior.Color = RGB(220, 240, 220)
        End If
        
        ' Update Named Ranges
        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("K3")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("K4")
        
        ' 1. Matach Manual Button
        Dim shpManual2 As Shape
        Set shpManual2 = wsMain.Shapes("shpMatachManual")
        If Not shpManual2 Is Nothing Then
            shpManual2.Left = wsMain.Range("K3").Left + (wsMain.Range("K3").Width - shpManual2.Width) / 2
            shpManual2.Top = wsMain.Range("K3").Top + 50
        End If
        
        wsMain.Range("A25:Z30").ClearContents
        
        Dim shpToggleLayout2 As Shape
        Set shpToggleLayout2 = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout2 Is Nothing Then
            shpToggleLayout2.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)
        End If
    End If
    
    wsMain.Protect Password:="1234", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
"""

content = content + "\n" + new_toggle_macro

# Update version
content = content.replace('APP_VERSION As String = "2.271"', 'APP_VERSION As String = "2.272"')
content = content.replace('VERSION: V2.271', 'VERSION: V2.272')
content = content.replace('Error in V2.271!', 'Error in V2.272!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_271"', 'Attribute VB_Name = "Goren_Claude_V2_272"')

changelog = """' CHANGES IN 2.272:
'   - UI: Fixed bug where "Operation Manual" button was deleted on open.
'   - UI: Refactored ToggleDisplayMode to cleanly Cut/Paste Matach table to avoid overlaps with Reset button.
"""
content = content.replace("' CHANGES IN 2.271:", changelog + "' CHANGES IN 2.271:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.272.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 272')
