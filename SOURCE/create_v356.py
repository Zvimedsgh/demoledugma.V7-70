import re

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.55.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update VB_Name
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_55"', 'Attribute VB_Name = "Goren_Claude_V3_56"')

new_end_slide = '''        ' Final Summary Slide (V3.56 - Stacked buttons, working Coffee slide)
        Dim sldEnd As Object
        Set sldEnd = ppPres.Slides.Add(ppPres.Slides.Count + 1, 12) ' 12=ppLayoutBlank
        
        ' Add a nice background gradient
        sldEnd.Background.Fill.TwoColorGradient Style:=1, Variant:=1 ' msoGradientHorizontal
        sldEnd.Background.Fill.ForeColor.RGB = RGB(240, 248, 255)
        sldEnd.Background.Fill.BackColor.RGB = RGB(200, 230, 255)

        ' Add title
        Dim shpTitle As Object
        Set shpTitle = sldEnd.Shapes.AddTextbox(1, 100, 20, 760, 100)
        With shpTitle.TextFrame.TextRange
            .Text = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & _
                    ChrW(1492) & ChrW(1493) & ChrW(1508) & ChrW(1511) & ChrW(1492) & " " & _
                    ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!"
            .ParagraphFormat.Alignment = 2 ' ppAlignCenter
            .Font.Name = "Assistant"
            .Font.Size = 36
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(0, 51, 102)
        End With

        ' Layout: 3 buttons stacked vertically
        ' Slide Width = 960, Button Width = 300 -> Left = 330
        ' Button Height = 55
        ' Top1 = 150, Top2 = 230, Top3 = 310

        ' Button 1: Start Presentation
        Dim btnStart As Object
        Set btnStart = sldEnd.Shapes.AddShape(5, 330, 150, 300, 55) ' 5 = msoShapeRoundedRectangle
        btnStart.Fill.ForeColor.RGB = RGB(0, 120, 215)
        btnStart.Line.Visible = msoFalse
        With btnStart.TextFrame.TextRange
            .Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
            .Font.Size = 20
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnStart.ActionSettings(1).Action = 3 ' ppActionFirstSlide
        btnStart.ActionSettings(1).SoundEffect.Name = "Click"

        ' Button 2: Close and Return
        Dim btnClose As Object
        Set btnClose = sldEnd.Shapes.AddShape(5, 330, 230, 300, 55)
        btnClose.Fill.ForeColor.RGB = RGB(215, 60, 60)
        btnClose.Line.Visible = msoFalse
        With btnClose.TextFrame.TextRange
            .Text = ChrW(1505) & ChrW(1490) & ChrW(1493) & ChrW(1512) & " " & ChrW(1493) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
            .Font.Size = 20
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnClose.ActionSettings(1).Action = 6 ' ppActionEndShow
        btnClose.ActionSettings(1).SoundEffect.Name = "Click"

        ' Button 3: Coffee
        Dim btnCoffee As Object
        Set btnCoffee = sldEnd.Shapes.AddShape(5, 330, 310, 300, 55)
        btnCoffee.Fill.ForeColor.RGB = RGB(255, 140, 0)
        btnCoffee.Line.Visible = msoFalse
        With btnCoffee.TextFrame.TextRange
            .Text = ChrW(1500) & ChrW(1498) & " " & ChrW(1500) & ChrW(1492) & ChrW(1499) & ChrW(1497) & ChrW(1503) & " " & ChrW(1511) & ChrW(1508) & ChrW(1492)
            .Font.Size = 20
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        ' Jump to the next slide (which will be the coffee slide)
        btnCoffee.ActionSettings(1).Action = 1 ' ppActionNextSlide
        btnCoffee.ActionSettings(1).SoundEffect.Name = "Click"
        
        ' ===============================
        ' Coffee Slide (Visible)
        ' ===============================
        Dim sldCoffee As Object
        Set sldCoffee = ppPres.Slides.Add(ppPres.Slides.Count + 1, 12)
        sldCoffee.Background.Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        ' Add Coffee Picture from SOURCE folder
        On Error Resume Next
        Dim coffeePath As String
        coffeePath = ThisWorkbook.Path & "\\coffee.jpg"
        If Dir(coffeePath) <> "" Then
            Dim picCoffee As Object
            Set picCoffee = sldCoffee.Shapes.AddPicture(coffeePath, 0, -1, 230, 20, 500, 500)
        End If
        On Error GoTo ERR_HANDLER
        
        ' Add Back Button
        Dim btnBack As Object
        Set btnBack = sldCoffee.Shapes.AddShape(5, 430, 450, 100, 50)
        btnBack.Fill.ForeColor.RGB = RGB(100, 100, 100)
        btnBack.Line.Visible = msoFalse
        With btnBack.TextFrame.TextRange
            .Text = ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) ' "חזור"
            .Font.Size = 20
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnBack.ActionSettings(1).Action = 2 ' ppActionPreviousSlide (goes back to Summary)
        btnBack.ActionSettings(1).SoundEffect.Name = "Click"
        ' ==============================='''

old_end_slide_start = "' Final Summary Slide"
old_end_slide_end = "' NOW show the presentation in Slide Show mode"

idx1 = content.find(old_end_slide_start)
idx2 = content.find(old_end_slide_end)

if idx1 != -1 and idx2 != -1:
    content = content[:idx1] + new_end_slide + '\n\n' + content[idx2:]
else:
    print("Could not find end slide boundaries")


with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.56.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V3.56 created successfully.")
