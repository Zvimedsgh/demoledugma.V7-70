with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# Fix the Clear bug that removes the green background
pattern1 = r'    \' Clear previous mess\n    wsMain\.Range\("B14:K15"\)\.Clear\n    With wsMain\.Range\("A2:B4"\)\n        \.UnMerge\n        \.Clear\n    End With'
replacement1 = """    ' Clear previous mess without removing background colors
    wsMain.Range("B14:K15").ClearContents
    wsMain.Range("B14:K15").Interior.Color = RGB(220, 240, 220)
    With wsMain.Range("A2:B4")
        .UnMerge
        .ClearContents
        .Interior.Color = RGB(220, 240, 220)
    End With"""
text = re.sub(pattern1, replacement1, text, flags=re.DOTALL)

# Change the shape fill color from light gray to prominent yellow
pattern2 = r'shpInstall\.Fill\.ForeColor\.RGB = RGB\(245, 245, 245\)'
replacement2 = r'shpInstall.Fill.ForeColor.RGB = RGB(255, 255, 153) \' Prominent pastel yellow'
text = re.sub(pattern2, replacement2, text)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated background and shape fill successfully.")
