import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.110.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("shpAgency.TextFrame.TextRange.ParagraphFormat.Alignment = 3", "shpAgency.TextFrame.TextRange.ParagraphFormat.Alignment = 1")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.110.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("Fixed alignment in V2.110.")
