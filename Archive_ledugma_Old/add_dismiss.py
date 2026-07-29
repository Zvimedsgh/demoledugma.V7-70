import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 1. Add creation of shpDismissMsg
    target_str = 'shpInstall.OnAction = "OpenInstallWeb"'
    if 'shpDismissMsg' not in content:
        dismiss_shape_code = """
    Dim shpDismiss As Shape
    On Error Resume Next
    Set shpDismiss = wsMain.Shapes("shpDismissMsg")
    On Error GoTo 0
    If shpDismiss Is Nothing Then
        Set shpDismiss = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, shpInstall.Left, shpInstall.Top + shpInstall.Height + 5, shpInstall.Width, 20)
        shpDismiss.Name = "shpDismissMsg"
    End If
    shpDismiss.Fill.ForeColor.RGB = RGB(240, 240, 240)
    shpDismiss.Line.Visible = msoFalse
    With shpDismiss.TextFrame2.TextRange
        .Text = ChrW(1492) & ChrW(1489) & ChrW(1504) & ChrW(1514) & ChrW(1497) & ChrW(44) & " " & ChrW(1488) & ChrW(1500) & " " & ChrW(1514) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & ChrW(1497) & ChrW(1493) & ChrW(1514) & ChrW(1512)
        .Font.Size = 9
        .Font.Fill.ForeColor.RGB = RGB(100, 100, 100)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    shpDismiss.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shpDismiss.OnAction = "DismissInstallMsg"
"""
        content = content.replace(target_str, target_str + "\n" + dismiss_shape_code)
    
    # 2. Add DismissInstallMsg macro
    if 'Public Sub DismissInstallMsg()' not in content:
        dismiss_macro = """
Public Sub DismissInstallMsg()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Shapes("shpInstallMsg").Delete
    ws.Shapes("shpDismissMsg").Delete
    On Error GoTo 0
End Sub
"""
        content += "\n" + dismiss_macro

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Success!")

if __name__ == "__main__":
    main()
