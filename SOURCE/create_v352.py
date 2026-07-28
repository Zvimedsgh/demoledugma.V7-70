import re

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.51.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update VB_Name
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_51"', 'Attribute VB_Name = "Goren_Claude_V3_52"')

# New end slide logic
new_end_slide = '''        ' Final Summary Slide (V3.52 - Buttons and Coffee)
        Dim sldEnd As Object
        Set sldEnd = ppPres.Slides.Add(ppPres.Slides.Count + 1, 12) ' 12=ppLayoutBlank
        
        ' Add a nice background gradient
        sldEnd.Background.Fill.TwoColorGradient Style:=1, Variant:=1 ' msoGradientHorizontal
        sldEnd.Background.Fill.ForeColor.RGB = RGB(240, 248, 255)
        sldEnd.Background.Fill.BackColor.RGB = RGB(200, 230, 255)

        ' Add title
        Dim shpTitle As Object
        Set shpTitle = sldEnd.Shapes.AddTextbox(1, 100, 30, 760, 100)
        With shpTitle.TextFrame.TextRange
            .Text = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & _
                    ChrW(1492) & ChrW(1493) & ChrW(1508) & ChrW(1511) & ChrW(1492) & " " & _
                    ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!"
            .ParagraphFormat.Alignment = 2 ' ppAlignCenter
            .Font.Name = "Assistant"
            .Font.Size = 44
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(0, 51, 102)
        End With

        ' Button 1: Start Presentation
        Dim btnStart As Object
        Set btnStart = sldEnd.Shapes.AddShape(166, 280, 180, 400, 70) ' 166 = msoShapeRoundRect
        btnStart.Fill.ForeColor.RGB = RGB(0, 120, 215)
        btnStart.Line.Visible = msoFalse
        With btnStart.TextFrame.TextRange
            .Text = "1. " & ChrW(1492) & ChrW(1510) & ChrW(1490) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
            .Font.Size = 24
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnStart.ActionSettings(1).Action = 3 ' ppActionFirstSlide
        btnStart.ActionSettings(1).SoundEffect.Name = "Click"

        ' Button 2: Close and Return
        Dim btnClose As Object
        Set btnClose = sldEnd.Shapes.AddShape(166, 280, 270, 400, 70)
        btnClose.Fill.ForeColor.RGB = RGB(215, 60, 60)
        btnClose.Line.Visible = msoFalse
        With btnClose.TextFrame.TextRange
            .Text = "2. " & ChrW(1505) & ChrW(1490) & ChrW(1493) & ChrW(1512) & " " & ChrW(1493) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
            .Font.Size = 24
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnClose.ActionSettings(1).Action = 6 ' ppActionEndShow

        ' Button 3: Coffee
        Dim btnCoffee As Object
        Set btnCoffee = sldEnd.Shapes.AddShape(166, 280, 360, 400, 70)
        btnCoffee.Fill.ForeColor.RGB = RGB(255, 140, 0)
        btnCoffee.Line.Visible = msoFalse
        With btnCoffee.TextFrame.TextRange
            .Text = "3. " & ChrW(1500) & ChrW(1498) & " " & ChrW(1500) & ChrW(1492) & ChrW(1499) & ChrW(1497) & ChrW(1503) & " " & ChrW(1511) & ChrW(1508) & ChrW(1492)
            .Font.Size = 24
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        
        ' Add Coffee Picture from SOURCE folder
        On Error Resume Next
        Dim coffeePath As String
        coffeePath = ThisWorkbook.Path & "\\coffee.jpg"
        If Dir(coffeePath) <> "" Then
            Dim picCoffee As Object
            Set picCoffee = sldEnd.Shapes.AddPicture(coffeePath, 0, -1, 440, 440, 80, 80) ' msoFalse, msoTrue
            ' Center it horizontally at the bottom
            ' btnCoffee is at Top 360. Pic is at Top 440
        End If
        On Error GoTo ERR_HANDLER'''

old_end_slide_start = "' Final Summary Slide"
old_end_slide_end = "        Dim pg As Long"

idx1 = content.find(old_end_slide_start)
idx2 = content.find(old_end_slide_end)

if idx1 != -1 and idx2 != -1:
    content = content[:idx1] + new_end_slide + '\n\n' + content[idx2:]
else:
    print("Could not find end slide boundaries")


# Now we also need to change how the presentation is opened at the very end
new_show_logic = '''
' NOW show the presentation in Slide Show mode on the last slide
On Error Resume Next
If Not ppPres Is Nothing Then
    ppApp.Visible = True
    AppActivate ppApp.Caption
    
    Dim ssw As Object
    Set ssw = ppPres.SlideShowSettings.Run
    ssw.View.GotoSlide ppPres.Slides.Count
End If
'''

old_show_logic_start = "' NOW show the presentation maximized and minimize its ribbon"
old_show_logic_end = "On Error GoTo ERR_HANDLER"

idx3 = content.find(old_show_logic_start)
idx4 = content.find(old_show_logic_end, idx3)

if idx3 != -1 and idx4 != -1:
    content = content[:idx3] + new_show_logic + '\n' + content[idx4:]
else:
    print("Could not find show logic boundaries")


with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.52.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V3.52 created successfully.")
