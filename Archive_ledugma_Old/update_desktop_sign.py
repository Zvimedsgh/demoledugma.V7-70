import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    line2_1 = ' & '.join(f'ChrW({ord(c)})' for c in "ניתן לשלוח הודעת ")
    line2_2 = ' & '.join(f'ChrW({ord(c)})' for c in "וואטסאפ ")
    line2_3 = ' & '.join(f'ChrW({ord(c)})' for c in "לטלפון")
    line3 = '"054-6677396"'
    line4 = ' & '.join(f'ChrW({ord(c)})' for c in "לקבלת עזרה")
    
    vba_str = f'{line2_1} & {line2_2} & {line2_3} & vbCrLf & _\n{line3} & vbCrLf & _\n{line4}'

    new_hl_code = f'''
    ' Desktop: Remove Hyperlink, OnAction, and set text to 3 lines
    Dim hl As Hyperlink
    On Error Resume Next
    For Each hl In wsMain.Hyperlinks
        If hl.Shape.Name = "shpInstallMsg" Then hl.Delete
    Next hl
    With wsMain.Shapes("shpInstallMsg")
        .OnAction = ""
        .TextFrame2.TextRange.Text = {vba_str}
    End With
    On Error GoTo 0
'''

    # Find the old Hyperlink delete code and replace it
    pattern = re.compile(r"    ' Delete Hyperlink from shpInstallMsg so Desktop uses OnAction\n.*?On Error GoTo 0\n", re.DOTALL)
    if pattern.search(content):
        content = pattern.sub(new_hl_code, content)
        print("Updated ApplyDemoLockOnOpen to 3 lines & no macro.")
    else:
        print("Could not find the previous hyperlink deletion code to replace.")

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
