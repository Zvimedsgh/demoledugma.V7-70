import re

with open(r'C:\LEVAV PROJECT\SOURCE\old_bp_clean.txt', 'r', encoding='utf-8') as f:
    bp_clean = f.read()

new_end_slide = '''        ' Final Summary Slide (V3.61 - Clean, no coffee button)
        Dim sldEnd As Object
        Set sldEnd = ppPres.Slides.Add(ppPres.Slides.Count + 1, 12) ' 12=ppLayoutBlank
        
        ' Add a nice background gradient
        sldEnd.Background.Fill.TwoColorGradient Style:=1, Variant:=1 ' msoGradientHorizontal
        sldEnd.Background.Fill.ForeColor.RGB = RGB(240, 248, 255)
        sldEnd.Background.Fill.BackColor.RGB = RGB(200, 230, 255)

        ' Add title
        Dim shpTitle As Object
        Set shpTitle = sldEnd.Shapes.AddTextbox(1, 100, 50, 760, 100)
        With shpTitle.TextFrame.TextRange
            .Text = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & _
                    ChrW(1492) & ChrW(1493) & ChrW(1508) & ChrW(1511) & ChrW(1492) & " " & _
                    ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!"
            .ParagraphFormat.Alignment = 2 ' ppAlignCenter
            .Font.Name = "Assistant"
            .Font.Size = 40
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(0, 51, 102)
        End With

        ' Layout: 2 buttons stacked vertically
        ' Slide Width = 960, Button Width = 300 -> Left = 330
        ' Button Height = 60
        ' Top1 = 200, Top2 = 290

        ' Button 1: Start Presentation
        Dim btnStart As Object
        Set btnStart = sldEnd.Shapes.AddShape(5, 330, 200, 300, 60) ' 5 = msoShapeRoundedRectangle
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
        Set btnClose = sldEnd.Shapes.AddShape(5, 330, 290, 300, 60)
        btnClose.Fill.ForeColor.RGB = RGB(215, 60, 60)
        btnClose.Line.Visible = msoFalse
        With btnClose.TextFrame.TextRange
            .Text = ChrW(1505) & ChrW(1490) & ChrW(1493) & ChrW(1512) & " " & ChrW(1493) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
            .Font.Size = 22
            .Font.Bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = 2
        End With
        btnClose.ActionSettings(1).Action = 6 ' ppActionEndShow
        btnClose.ActionSettings(1).SoundEffect.Name = "Click"
'''

idx_start = bp_clean.find("' Final Summary Slide")
idx_end = bp_clean.find("        Dim pg As Long")

if idx_start != -1 and idx_end != -1:
    bp_clean = bp_clean[:idx_start] + new_end_slide + '\n' + bp_clean[idx_end:]
else:
    print("Failed to find boundaries in old_bp_clean.txt")
    exit(1)


# Now we have a perfect BuildPresentation.
# We need to inject it into V3.60 to replace its broken BuildPresentation.
with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.60.bas', 'r', encoding='utf-8') as f:
    v360 = f.read()

# Update VB_Name
v360 = v360.replace('Attribute VB_Name = "Goren_Claude_V3_60"', 'Attribute VB_Name = "Goren_Claude_V3_61"')

# Find boundaries of BuildPresentation in V3.60
bp_start = v360.find("Public Sub BuildPresentation()")
bp_end = v360.find("End Sub\n\n' ============================================================================\n' HELPER: Robust Chart Export")

if bp_start != -1 and bp_end != -1:
    # Replace the broken BuildPresentation with the perfect one
    v361 = v360[:bp_start] + bp_clean + v360[bp_end + 8:]
else:
    print("Failed to find boundaries in V3.60")
    exit(1)

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.61.bas', 'w', encoding='utf-8') as f:
    f.write(v361)
print("V3.61 created")
