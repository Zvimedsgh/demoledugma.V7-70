import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.109.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Rename AddPageNumber to AddSlideFooter and update its signature
content = content.replace(
    "Private Sub AddPageNumber(ByVal ppSlide As Object, ByVal pageNum As Long, ByVal totalPages As Long, ByVal slideW As Single, ByVal slideH As Single)",
    "Private Sub AddSlideFooter(ByVal ppSlide As Object, ByVal pageNum As Long, ByVal totalPages As Long, ByVal slideW As Single, ByVal slideH As Single, ByVal agencyName As String)"
)

# 2. Add the agency name shape into AddSlideFooter
footer_code = """    Dim shpAgency As Object
    Set shpAgency = ppSlide.Shapes.AddTextbox(1, 20, slideH - 22, slideW - 140, 20)
    shpAgency.TextFrame.TextRange.Text = agencyName
    shpAgency.TextFrame.TextRange.Font.Name = "Arial"
    shpAgency.TextFrame.TextRange.Font.Size = 10
    shpAgency.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
    shpAgency.TextFrame.TextRange.ParagraphFormat.Alignment = 1 ' Left align (or Right for Hebrew, let's use Right which is 3)
    shpAgency.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    shpAgency.TextFrame.WordWrap = False
    shpAgency.TextFrame.MarginTop = 0
    shpAgency.TextFrame.MarginBottom = 0
    shpAgency.ZOrder 0
"""
content = content.replace("    shp.ZOrder 0\n", "    shp.ZOrder 0\n\n" + footer_code)

# 3. Update caller loop in BuildPresentation
old_caller = "        AddPageNumber ppPres.Slides(pg), pg, ppPres.Slides.Count, slideW, slideH"
new_caller = "        AddSlideFooter ppPres.Slides(pg), pg, ppPres.Slides.Count, slideW, slideH, GetActiveAgencyName()"
content = content.replace(old_caller, new_caller)

# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_109"', 'Attribute VB_Name = "Goren_Claude_V2_110"')
content = content.replace('VERSION: V2.109', 'VERSION: V2.110')
content = content.replace('APP_VERSION As String = "2.109"', 'APP_VERSION As String = "2.110"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.110.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.110 created.")
