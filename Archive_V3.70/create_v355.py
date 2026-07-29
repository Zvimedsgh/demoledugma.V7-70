import re

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.54.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update VB_Name
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_54"', 'Attribute VB_Name = "Goren_Claude_V3_55"')

new_end_slide = '''        ' Final Summary Slide (V3.55 - Beautiful buttons & Coffee Slide)
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
        ' Slide Width = 960. 3 buttons of width 200 -> 600 total width.
        ' Gap between buttons = (960 - 600) / 4 = 90
        ' Left1=90, Left2=380, Left3=670

        ' Button 1: Start Presentation
        Dim btnStart As Object
        Set btnStart = sldEnd.Shapes.AddShape(5, 90, 240, 200, 80) ' 5 = msoShapeRoundedRectangle
        btnStart.Fill.ForeColor.RGB = RGB(0, 120, 215)
        btnStart.Line.Visible = msoFalse
        With btnStart.TextFrame.TextRange
            .Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
            .Font.Size = 22
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnStart.ActionSettings(1).Action = 3 ' ppActionFirstSlide
        btnStart.ActionSettings(1).SoundEffect.Name = "Click"

        ' Button 2: Close and Return
        Dim btnClose As Object
        Set btnClose = sldEnd.Shapes.AddShape(5, 380, 240, 200, 80)
        btnClose.Fill.ForeColor.RGB = RGB(215, 60, 60)
        btnClose.Line.Visible = msoFalse
        With btnClose.TextFrame.TextRange
            .Text = ChrW(1505) & ChrW(1490) & ChrW(1493) & ChrW(1512) & vbCrLf & ChrW(1493) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
            .Font.Size = 20
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnClose.ActionSettings(1).Action = 6 ' ppActionEndShow
        btnClose.ActionSettings(1).SoundEffect.Name = "Click"

        ' Button 3: Coffee
        Dim btnCoffee As Object
        Set btnCoffee = sldEnd.Shapes.AddShape(5, 670, 240, 200, 80)
        btnCoffee.Fill.ForeColor.RGB = RGB(255, 140, 0)
        btnCoffee.Line.Visible = msoFalse
        With btnCoffee.TextFrame.TextRange
            .Text = ChrW(1500) & ChrW(1498) & " " & ChrW(1500) & ChrW(1492) & ChrW(1499) & ChrW(1497) & ChrW(1503) & " " & ChrW(1511) & ChrW(1508) & ChrW(1492)
            .Font.Size = 22
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnCoffee.ActionSettings(1).Action = 4 ' ppActionLastSlide (Goes to Coffee slide)
        btnCoffee.ActionSettings(1).SoundEffect.Name = "Click"
        
        ' ===============================
        ' Coffee Slide (Hidden)
        ' ===============================
        Dim sldCoffee As Object
        Set sldCoffee = ppPres.Slides.Add(ppPres.Slides.Count + 1, 12)
        ' Hide this slide so it doesn't show during normal scrolling
        sldCoffee.SlideShowTransition.Hidden = msoTrue
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
        btnBack.ActionSettings(1).Action = 5 ' ppActionLastSlideViewed
        btnBack.ActionSettings(1).SoundEffect.Name = "Click"
        ' ===============================
'''

old_end_slide_start = "' Final Summary Slide"
old_end_slide_end = "        Dim pg As Long"

idx1 = content.find(old_end_slide_start)
idx2 = content.find(old_end_slide_end)

if idx1 != -1 and idx2 != -1:
    content = content[:idx1] + new_end_slide + '\n\n' + content[idx2:]
else:
    print("Could not find end slide boundaries")

# Update slideShow GotoSlide to go to the Summary slide, not the Coffee slide!
# Since Coffee slide is the last one (Slides.Count), the summary slide is Slides.Count - 1
new_show_logic = '''
' NOW show the presentation in Slide Show mode on the summary slide
On Error Resume Next
If Not ppPres Is Nothing Then
    ppApp.Visible = True
    AppActivate ppApp.Caption
    
    Dim ssw As Object
    Set ssw = ppPres.SlideShowSettings.Run
    ssw.View.GotoSlide ppPres.Slides.Count - 1
End If
'''

old_show_logic_start = "' NOW show the presentation in Slide Show mode on the last slide"
old_show_logic_end = "End If"

idx3 = content.find(old_show_logic_start)
idx4 = content.find(old_show_logic_end, idx3)

if idx3 != -1 and idx4 != -1:
    content = content[:idx3] + new_show_logic + content[idx4 + 6:]
else:
    print("Could not find show logic boundaries")


with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.55.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V3.55 created successfully.")
