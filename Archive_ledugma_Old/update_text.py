import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.274.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace A16 with A19 for the version string
content = content.replace('wsMain.Range("A16").Value = ', 'wsMain.Range("A19").Value = ')
content = content.replace('wsMain.Range("A16").Font', 'wsMain.Range("A19").Font')
content = content.replace('wsMain.Range("A16").HorizontalAlignment', 'wsMain.Range("A19").HorizontalAlignment')

# For the credit text, replace A24:L24 with G19
content = content.replace('wsMain.Range("A24:L24")', 'wsMain.Range("G19")')
# Remove the MergeCells = True line inside the With block
content = content.replace('.MergeCells = True', '')

content = content.replace('APP_VERSION As String = "2.274"', 'APP_VERSION As String = "2.275"')
content = content.replace('VERSION: V2.274', 'VERSION: V2.275')
content = content.replace('Error in V2.274!', 'Error in V2.275!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_274"', 'Attribute VB_Name = "Goren_Claude_V2_275"')

changelog = """' CHANGES IN 2.275:
'   - UI: Moved version number to A19 and Credit text to G19.
"""
content = content.replace("' CHANGES IN 2.274:", changelog + "' CHANGES IN 2.274:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.275.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 275')
