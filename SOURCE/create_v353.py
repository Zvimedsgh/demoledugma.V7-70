import re

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.52.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update VB_Name
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_52"', 'Attribute VB_Name = "Goren_Claude_V3_53"')

new_end_slide = '''        ' Final Summary Slide (V3.53 - Narrower, taller buttons side by side)
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

        ' Layout: 3 buttons side-by-side
        ' Slide Width = 960. 3 buttons of width 250 -> 750 total width.
        ' Gap between buttons = (960 - 750) / 4 = 52.5
        ' B1 Left = 52.5
        ' B2 Left = 52.5*2 + 250 = 355
        ' B3 Left = 52.5*3 + 500 = 657.5
        ' Let's use: Left1=70, Left2=355, Left3=640

        ' Button 1: Start Presentation
        Dim btnStart As Object
        Set btnStart = sldEnd.Shapes.AddShape(166, 70, 220, 250, 140) ' 166 = msoShapeRoundRect
        btnStart.Fill.ForeColor.RGB = RGB(0, 120, 215)
        btnStart.Line.Visible = msoFalse
        With btnStart.TextFrame.TextRange
            .Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
            .Font.Size = 28
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnStart.ActionSettings(1).Action = 3 ' ppActionFirstSlide
        btnStart.ActionSettings(1).SoundEffect.Name = "Click"

        ' Button 2: Close and Return
        Dim btnClose As Object
        Set btnClose = sldEnd.Shapes.AddShape(166, 355, 220, 250, 140)
        btnClose.Fill.ForeColor.RGB = RGB(215, 60, 60)
        btnClose.Line.Visible = msoFalse
        With btnClose.TextFrame.TextRange
            .Text = ChrW(1505) & ChrW(1490) & ChrW(1493) & ChrW(1512) & vbCrLf & ChrW(1493) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
            .Font.Size = 28
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnClose.ActionSettings(1).Action = 6 ' ppActionEndShow
        btnClose.ActionSettings(1).SoundEffect.Name = "Click"

        ' Button 3: Coffee
        Dim btnCoffee As Object
        Set btnCoffee = sldEnd.Shapes.AddShape(166, 640, 220, 250, 140)
        btnCoffee.Fill.ForeColor.RGB = RGB(255, 140, 0)
        btnCoffee.Line.Visible = msoFalse
        With btnCoffee.TextFrame.TextRange
            .Text = ChrW(1500) & ChrW(1498) & " " & ChrW(1500) & ChrW(1492) & ChrW(1499) & ChrW(1497) & ChrW(1503) & " " & ChrW(1511) & ChrW(1508) & ChrW(1492)
            .Font.Size = 28
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        ' To make it clickable and show hand cursor without doing much:
        ' Link it to a dummy website, or use NextSlide (which does nothing on the last slide)
        btnCoffee.ActionSettings(1).Action = 2 ' ppActionNextSlide (does nothing since it's the last slide, but shows a hand cursor)
        btnCoffee.ActionSettings(1).SoundEffect.Name = "Click"
        
        ' Add Coffee Picture from SOURCE folder
        On Error Resume Next
        Dim coffeePath As String
        coffeePath = ThisWorkbook.Path & "\\coffee.jpg"
        If Dir(coffeePath) <> "" Then
            Dim picCoffee As Object
            ' Place it right below button 3: Top = 370, Left = 640 + 250/2 - 50 = 715
            Set picCoffee = sldEnd.Shapes.AddPicture(coffeePath, 0, -1, 715, 380, 100, 100)
            ' Make picture clickable too!
            picCoffee.ActionSettings(1).Action = 2
            picCoffee.ActionSettings(1).SoundEffect.Name = "Click"
        End If
        On Error GoTo ERR_HANDLER'''

# We need to replace the entire old end slide logic up to ' Add page numbers to all slides
old_end_slide_start = "' Final Summary Slide"
old_end_slide_end = "        Dim pg As Long"

idx1 = content.find(old_end_slide_start)
idx2 = content.find(old_end_slide_end)

if idx1 != -1 and idx2 != -1:
    content = content[:idx1] + new_end_slide + '\n\n' + content[idx2:]
else:
    print("Could not find end slide boundaries")

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.53.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V3.53 created successfully.")
