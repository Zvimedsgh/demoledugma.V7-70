import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.176.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.177.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(len(lines)):
    # 1. Move Demo Msg from E19:K20 to D19:K20
    if '"E19:K20"' in lines[i]:
        lines[i] = lines[i].replace('"E19:K20"', '"D19:K20"')
    if '"E19"' in lines[i]:
        lines[i] = lines[i].replace('"E19"', '"D19"')
        
    # 2. Right-align the 7 buttons
    if 'shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter' in lines[i]:
        # Only for the main buttons block (lines 3800-4000)
        if 3800 < i < 4000:
            lines[i] = lines[i].replace('msoAlignCenter', 'msoAlignRight')
            # Add MarginRight to give it a little breathing room from the edge
            lines.insert(i+1, lines[i].replace('.ParagraphFormat.Alignment = msoAlignRight', '.MarginRight = 10'))

# 3. Update version string
content = "".join(lines)
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_176"', 'Attribute VB_Name = "Goren_Claude_V2_177"')
content = content.replace('VERSION: V2.176', 'VERSION: V2.177')
content = content.replace('APP_VERSION As String = "2.176"', 'APP_VERSION As String = "2.177"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.177 correctly!")
