import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.224.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # We need to replace the With wsMain.Range("D19:K20") block in A00_SetupMainSheet
    # AND the one in ApplyDemoLockOnOpen
    
    line1_part1 = ' & '.join(f'ChrW({ord(c)})' for c in "לראות גיליון הוראות")
    line1_part2 = ' & '.join(f'ChrW({ord(c)})' for c in " ה קלק כאן ")
    line1_part3 = ' & '.join(f'ChrW({ord(c)})' for c in "קליק ימני פ")
    line2 = '"054-6677396"'
    line3 = ' & '.join(f'ChrW({ord(c)})' for c in "שלח וואטסאפ לעזרה")
    vba_str = f'{line1_part1} & {line1_part2} & {line1_part3} & vbCrLf & _\n{line2} & vbCrLf & _\n{line3}'

    new_setup_shape = """    ' Clear previous mess
    wsMain.Range("B14:K15").Clear
    wsMain.Range("D19:K20").Clear
    On Error Resume Next
    wsMain.Shapes("shpInstallMsg").Delete
    On Error GoTo ERR_HANDLER
    
    Dim shpInstall As Shape
    Set shpInstall = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("A2").Left + 10, wsMain.Range("A2").Top, 220, 65)
    shpInstall.Name = "shpInstallMsg"
    shpInstall.Fill.ForeColor.RGB = RGB(245, 245, 245)
    shpInstall.Line.ForeColor.RGB = RGB(0, 0, 139)
    shpInstall.Line.Weight = 2
    
    With shpInstall.TextFrame2.TextRange
        .Text = """ + vba_str + """
        .Font.Size = 12
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(0, 0, 139)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    shpInstall.TextFrame2.VerticalAnchor = msoAnchorMiddle
    
    On Error Resume Next
    wsMain.Hyperlinks.Add Anchor:=shpInstall, Address:="https://gorentec-my.sharepoint.com/:b:/g/personal/zvi_gorentech_co_il/IQBt2Ms0oWLySpRBl-6OY68MAXnBN2jyzpD3y43ZHmIUBwU?e=y10XmN"
    On Error GoTo ERR_HANDLER"""

    # We do a targeted replace for A00_SetupMainSheet
    pattern_setup = re.compile(r'(    With wsMain\.Range\("D19:K20"\)\n        \.Merge\n.*?        \.Font\.Color = RGB\(0, 0, 255\)\n    End With)', flags=re.DOTALL)
    
    # We expect TWO matches. 
    # Match 1: SetupMainSheet (line 3835)
    # Match 2: ApplyDemoLockOnOpen (line 7810)
    
    matches = pattern_setup.findall(content)
    if len(matches) == 2:
        # Replace first match with new_setup_shape
        content = content.replace(matches[0], new_setup_shape, 1)
        # Replace second match with new_setup_shape
        content = content.replace(matches[1], new_setup_shape, 1)
        print("Replaced both occurrences.")
    else:
        print(f"Found {len(matches)} occurrences, expected 2.")

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
