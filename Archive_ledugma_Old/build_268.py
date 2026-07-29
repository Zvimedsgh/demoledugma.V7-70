import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.267.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.267"', 'APP_VERSION As String = "2.268"')
content = content.replace('VERSION: V2.267', 'VERSION: V2.268')
content = content.replace('Error in V2.267!', 'Error in V2.268!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_267"', 'Attribute VB_Name = "Goren_Claude_V2_268"')

# Changelog
changelog = """' CHANGES IN 2.268:
'   - BUGFIX: Removed hallucinated 'Call get_rates' and replaced with proper value copying between Desktop and Laptop modes.
"""
content = content.replace("' CHANGES IN 2.267:", changelog + "' CHANGES IN 2.267:")

# Fix the macro
old_macro = """        Dim shpToggleLayout As Shape
        Set shpToggleLayout = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout Is Nothing Then
            shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1512) & ChrW(1495) & ChrW(1489)
        End If
        
        Call get_rates
        
    Else
        ' Move from Laptop to Desktop
        wsMain.Range("F12:G14").ClearContents"""

new_macro = """        Dim shpToggleLayout As Shape
        Set shpToggleLayout = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout Is Nothing Then
            shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1512) & ChrW(1495) & ChrW(1489)
        End If
        
    Else
        ' Move from Laptop to Desktop
        Dim curUSD2 As Variant, curEUR2 As Variant
        curUSD2 = wsMain.Range("G13").Value
        curEUR2 = wsMain.Range("G14").Value

        wsMain.Range("F12:G14").ClearContents"""

content = content.replace(old_macro, new_macro)


old_macro_2 = """    If isDesktop Then
        ' Move from Desktop to Laptop
        wsMain.Range("J2:K4").ClearContents"""

new_macro_2 = """    If isDesktop Then
        ' Move from Desktop to Laptop
        Dim curUSD As Variant, curEUR As Variant
        curUSD = wsMain.Range("K3").Value
        curEUR = wsMain.Range("K4").Value

        wsMain.Range("J2:K4").ClearContents"""

content = content.replace(old_macro_2, new_macro_2)

old_macro_3 = """        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("G13")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("G14")"""

new_macro_3 = """        wsMain.Range("G13").Value = curUSD
        wsMain.Range("G14").Value = curEUR
        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("G13")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("G14")"""

content = content.replace(old_macro_3, new_macro_3)

old_macro_4 = """        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("K3")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("K4")"""

new_macro_4 = """        wsMain.Range("K3").Value = curUSD2
        wsMain.Range("K4").Value = curEUR2
        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("K3")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("K4")"""

content = content.replace(old_macro_4, new_macro_4)


old_macro_5 = """        Dim shpToggleLayout2 As Shape
        Set shpToggleLayout2 = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout2 Is Nothing Then
            shpToggleLayout2.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)
        End If
        
        Call get_rates
    End If"""

new_macro_5 = """        Dim shpToggleLayout2 As Shape
        Set shpToggleLayout2 = wsMain.Shapes("shpToggleLayout")
        If Not shpToggleLayout2 Is Nothing Then
            shpToggleLayout2.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)
        End If
        
    End If"""

content = content.replace(old_macro_5, new_macro_5)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.268.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 268')
