import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.137.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Fix SearchClientName Activate bug
target_search_activate = """' Activate the search sheet and put cursor in search cell
wsSearch.Activate"""

new_search_activate = """' Activate the search sheet and put cursor in search cell
wsSearch.Visible = xlSheetVisible
wsSearch.Activate"""

content = content.replace(target_search_activate, new_search_activate)

# 2. Fix Demo Message RTL rendering
# Need to insert TextDirection after Alignment
target_demo_align = "shpDemoMsg.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 ' msoAlignCenter"
new_demo_align = """shpDemoMsg.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 ' msoAlignCenter
shpDemoMsg.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2 ' msoTextDirectionRightToLeft"""

content = content.replace(target_demo_align, new_demo_align)

# Update Version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_137"', 'Attribute VB_Name = "Goren_Claude_V2_138"')
content = content.replace('VERSION: V2.137', 'VERSION: V2.138')
content = content.replace('APP_VERSION As String = "2.137"', 'APP_VERSION As String = "2.138"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.138 created.")

