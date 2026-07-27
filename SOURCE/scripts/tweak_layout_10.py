import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.29.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Skip wsMain in the font size 14 loop at the end
    old_loop = '3200  For Each wsLoop In ThisWorkbook.Worksheets\n3210  wsLoop.DisplayRightToLeft = True\n3220  wsLoop.Cells.Font.Size = 14\n3230  Next wsLoop'
    new_loop = '3200  For Each wsLoop In ThisWorkbook.Worksheets\n3210  wsLoop.DisplayRightToLeft = True\n3220  If wsLoop.Name <> CONTROL_SHEET_NAME() Then wsLoop.Cells.Font.Size = 14\n3230  Next wsLoop'
    content = content.replace(old_loop, new_loop)

    # 2. Set wsMain font size 14 at the very beginning of its formatting (so it gets overridden by specific fonts later)
    old_clear = 'wsMain.Range("A2:K20").ClearContents'
    new_clear = 'wsMain.Range("A2:K20").ClearContents\n    wsMain.Cells.Font.Size = 14'
    content = content.replace(old_clear, new_clear)

    # Bump version
    content = content.replace('Attribute VB_Name = "Goren_Claude1.29"', 'Attribute VB_Name = "Goren_Claude1.30"')
    content = content.replace("' VERSION: V1.29", "' VERSION: V1.30")
    content = content.replace('Private Const APP_VERSION As String = "1.29"', 'Private Const APP_VERSION As String = "1.30"')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
