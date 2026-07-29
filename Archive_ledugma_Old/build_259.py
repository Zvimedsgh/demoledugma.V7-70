import os
import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.258.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.258"', 'APP_VERSION As String = "2.259"')
content = content.replace('VERSION: V2.258', 'VERSION: V2.259')
content = content.replace('Error in V2.258!', 'Error in V2.259!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_258"', 'Attribute VB_Name = "Goren_Claude_V2_259"')

# Update Changelog
changelog = """' CHANGES IN 2.259:
'   - UI: Moved MATACH table and buttons to visible area (F12:I14, A12, A14) to solve laptop screen cutoff on the left side.
"""
content = content.replace("' CHANGES IN 2.258:", changelog + "' CHANGES IN 2.258:")

# Replacements for MATACH table
content = content.replace('Range("J2")', 'Range("F12")')
content = content.replace('Range("K2")', 'Range("G12")')
content = content.replace('Range("J2:K2")', 'Range("F12:G12")')
content = content.replace('Range("J3")', 'Range("F13")')
content = content.replace('Range("K3")', 'Range("G13")')
content = content.replace('Range("J4")', 'Range("F14")')
content = content.replace('Range("K4")', 'Range("G14")')
content = content.replace('Range("J3:J4")', 'Range("F13:F14")')
content = content.replace('Range("K3:K4")', 'Range("G13:G14")')
content = content.replace('Range("J2:K4")', 'Range("F12:G14")')

# Replacements for Status message
content = content.replace('Range("J5:L7")', 'Range("H12:J14")')
content = content.replace('Range("J5:K7")', 'Range("H12:I14")')
content = content.replace('Range("J5")', 'Range("H12")')

# Replacements for Buttons
content = re.sub(r'Set shpManual = wsMain\.Shapes\.AddShape\(msoShapeOval, wsMain\.Range\("L1"\)\.Left \+ \(wsMain\.Range\("L1"\)\.Width - 120\) / 2, wsMain\.Range\("L1"\)\.Top \+ wsMain\.Range\("L1"\)\.Height - 80, 120, 80\)',
                 r'Set shpManual = wsMain.Shapes.AddShape(msoShapeOval, wsMain.Range("A12").Left + 10, wsMain.Range("A12").Top, 100, 40)', content)

content = re.sub(r'Set shp = wsMain\.Shapes\.AddShape\(msoShapeRoundedRectangle, wsMain\.Range\("L19"\)\.Left, wsMain\.Range\("L19"\)\.Top, 160, 30\)',
                 r'Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("A14").Left + 10, wsMain.Range("A14").Top, 140, 30)', content)

# Adjust credit in case we used C16 instead of D16 in V2.258
content = content.replace('With wsMain.Range("C16:I16")', 'With wsMain.Range("D16:I16")')
content = content.replace('wsMain.Range("C16:I16").UnMerge', 'wsMain.Range("D16:I16").UnMerge')
# Wait, I didn't replace C16:I16 unmerge in V2.258 script, I only replaced With statement.
# Let's just blindly change C16:I16 to D16:I16 for safety.
content = content.replace('C16:I16', 'D16:I16')


with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.259.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 259')
