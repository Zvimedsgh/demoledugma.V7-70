import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.25.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Fix btnReset position to be perfectly centered between btnAll and btnSearch
    old_reset = 'Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G11").Left + (wsMain.Range("G11").Width + wsMain.Range("F11").Width - 120) / 2, wsMain.Range("F11").Top + 35, 120, 25)'
    new_reset = 'Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnAll.Left + (btnSearch.Left + btnSearch.Width - btnAll.Left - 120) / 2, wsMain.Range("F11").Top + 32, 120, 25)'
    content = content.replace(old_reset, new_reset)

    # 2. Fix the duplicated credit text.
    old_credit_app = 'If Len(creditVersionCell) > 0 Then\n        wsMain.Range(creditVersionCell).Value = creditText & " | v" & APP_VERSION\n        wsMain.Range(creditVersionCell).Font.Size = 9\n        wsMain.Range(creditVersionCell).Font.Color = RGB(0, 0, 102)\n        wsMain.Range(creditVersionCell).Font.Bold = True\n        wsMain.Range(creditVersionCell).VerticalAlignment = xlVAlignCenter\n    End If'
    new_credit_app = 'If Len(creditVersionCell) > 0 And creditVersionCell <> "A23" Then\n        On Error Resume Next\n        wsMain.Range(creditVersionCell).ClearContents\n        Err.Clear\n    End If\n    wsMain.Range("A23").Value = creditText & " - v" & APP_VERSION\n    wsMain.Range("A23").Font.Size = 10\n    wsMain.Range("A23").Font.Color = RGB(150, 150, 150)\n    wsMain.Range("A23").Font.Bold = False\n    wsMain.Range("A23").HorizontalAlignment = xlLeft'
    content = content.replace(old_credit_app, new_credit_app)

    # Bump version
    content = content.replace('Attribute VB_Name = "Goren_Claude1.25"', 'Attribute VB_Name = "Goren_Claude1.26"')
    content = content.replace("' VERSION: V1.25", "' VERSION: V1.26")
    content = content.replace('Private Const APP_VERSION As String = "1.25"', 'Private Const APP_VERSION As String = "1.26"')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
