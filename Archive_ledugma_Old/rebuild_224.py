import re
import datetime

def main():
    src_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.221.bas'
    dest_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.224.bas'
    
    with open(src_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Bump version
    content = content.replace('VERSION: V2.221', 'VERSION: V2.224')
    content = content.replace('Attribute VB_Name = "Goren_Claude_V2_221"', 'Attribute VB_Name = "Goren_Claude_V2_224"')
    
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Replace DATE: ... with new date and LATEST FIXES
    content = re.sub(r"' DATE: .*?\r?\n", f"' DATE: {now_str} (LATEST FIXES)\n", content, 1)

    # 2. F3 to F10
    content = content.replace('Application.Goto wsMain.Range("F3")', 'Application.Goto wsMain.Range("F10")')
    content = content.replace('Application.Goto wsMainUI.Range("F3")', 'Application.Goto wsMainUI.Range("F10")')
    content = content.replace('Application.Goto ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("F3")', 'Application.Goto ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("F10")')

    # 3. Add the Hebrew Shape instead of A2:B4 cells
    # We find where HideInstallationInstructions is, and modify it so it doesn't clear A2:B4 but instead deletes shpInstallMsg
    
    # In V2.221, A00_SetupMainSheet had this:
    old_setup_text = """    ' Add the 3-line installation message in A2:B4
    With wsMain.Range("A2:B4")
        .Merge
        .Value = ChrW(1512) & ChrW(1488) & ChrW(1492) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & vbCrLf & _
                 ChrW(1500) & ChrW(1506) & ChrW(1494) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & vbCrLf & _
                 ChrW(1513) & ChrW(1500) & ChrW(1495) & " " & "W" & "h" & "a" & "t" & "s" & "A" & "p" & "p" & " " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & " " & "0" & "5" & "4" & "-" & "6" & "6" & "7" & "7" & "3" & "9" & "6"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 12
        .Font.Bold = msoTrue
        .Font.Color = RGB(0, 0, 139)
    End With"""
    
    # We replace it with the Shape creation and Add Hyperlink
    line1_part1 = ' & '.join(f'ChrW({ord(c)})' for c in "לראות גיליון הוראות")
    line1_part2 = ' & '.join(f'ChrW({ord(c)})' for c in " ה קלק כאן ")
    line1_part3 = ' & '.join(f'ChrW({ord(c)})' for c in "קליק ימני פ")
    line2 = '"054-6677396"'
    line3 = ' & '.join(f'ChrW({ord(c)})' for c in "שלח וואטסאפ לעזרה")
    
    vba_str = f'{line1_part1} & {line1_part2} & {line1_part3} & vbCrLf & _\n{line2} & vbCrLf & _\n{line3}'

    new_setup_shape = """    ' Clear previous mess
    wsMain.Range("B14:K15").Clear
    With wsMain.Range("A2:B4")
        .UnMerge
        .Clear
    End With
    
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

    if old_setup_text in content:
        content = content.replace(old_setup_text, new_setup_shape)
    else:
        print("WARNING: Could not find old_setup_text")

    # 4. Remove Hyperlink in ApplyDemoLockOnOpen
    hl_remove_code = """
    ' Delete Hyperlink from shpInstallMsg so Desktop uses OnAction
    Dim hl As Hyperlink
    On Error Resume Next
    For Each hl In wsMain.Hyperlinks
        If hl.Shape.Name = "shpInstallMsg" Then hl.Delete
    Next hl
    On Error GoTo 0
"""
    old_apply_demo = """Public Sub ApplyDemoLockOnOpen()
    Dim wsMain As Worksheet
    On Error Resume Next
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If wsMain Is Nothing Then Exit Sub"""
    
    new_apply_demo = old_apply_demo + "\n" + hl_remove_code
    
    content = content.replace(old_apply_demo, new_apply_demo)
    
    # 5. Fix HideInstallationInstructions to not look for A2:B4 merge
    old_hide = """    ' Clear the manual text from main screen
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If Not wsMain Is Nothing Then
        With wsMain.Range("A2:B4")
            .UnMerge
            .ClearContents
            .Borders.LineStyle = xlNone
            .Interior.ColorIndex = xlNone
        End With
    End If
    On Error GoTo 0"""
    
    new_hide = """    ' Clear the manual text from main screen
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If Not wsMain Is Nothing Then
        wsMain.Shapes("shpInstallMsg").Delete
    End If
    On Error GoTo 0"""
    
    content = content.replace(old_hide, new_hide)

    with open(dest_file, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Created {dest_file}")

if __name__ == "__main__":
    main()
