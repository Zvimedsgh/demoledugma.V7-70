import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Update shpInstallMsg text to the 4 full lines in pure Hebrew
    line1 = ' & '.join(f'ChrW({ord(c)})' for c in "הקש כאן לפתיחת הוראות התקנה")
    line2_1 = ' & '.join(f'ChrW({ord(c)})' for c in "ניתן לשלוח הודעת ")
    line2_2 = ' & '.join(f'ChrW({ord(c)})' for c in "וואטסאפ ")
    line2_3 = ' & '.join(f'ChrW({ord(c)})' for c in "לטלפון")
    line3 = '"054-6677396"'
    line4 = ' & '.join(f'ChrW({ord(c)})' for c in "לקבלת עזרה")
    
    vba_str = f'{line1} & vbCrLf & _\n{line2_1} & {line2_2} & {line2_3} & vbCrLf & _\n{line3} & vbCrLf & _\n{line4}'
    
    pattern_install = re.compile(r'(    With shpInstall\.TextFrame2\.TextRange\n        \.Text = ).*?(\n        \.Font\.Size = 12)', re.DOTALL)
    if pattern_install.search(content):
        content = pattern_install.sub(r'\1' + vba_str + r'\2', content)

    # 2. Add Hyperlink to the shape right after OnAction in A00_SetupMainSheet
    hl_code = '\n    On Error Resume Next\n    wsMain.Hyperlinks.Add Anchor:=shpInstall, Address:="https://gorentec-my.sharepoint.com/:b:/g/personal/zvi_gorentech_co_il/IQBt2Ms0oWLySpRBl-6OY68MAXnBN2jyzpD3y43ZHmIUBwU?e=y10XmN"\n    On Error GoTo 0\n'
    if 'wsMain.Hyperlinks.Add Anchor:=shpInstall' not in content:
        content = content.replace('shpInstall.OnAction = "OpenInstructions"', 'shpInstall.OnAction = "OpenInstructions"' + hl_code)

    # 3. In ApplyDemoLockOnOpen, remove the hyperlink!
    hl_remove_code = '''
    ' Delete Hyperlink from shpInstallMsg so Desktop uses OnAction
    Dim hl As Hyperlink
    On Error Resume Next
    For Each hl In wsMain.Hyperlinks
        If hl.Shape.Name = "shpInstallMsg" Then hl.Delete
    Next hl
    On Error GoTo 0
'''
    if 'Delete Hyperlink from shpInstallMsg' not in content:
        # We find ApplyDemoLockOnOpen
        pattern_apply = re.compile(r'(Public Sub ApplyDemoLockOnOpen\(\)\n    Dim wsMain As Worksheet\n    On Error Resume Next\n    Set wsMain = ThisWorkbook\.Worksheets\(CONTROL_SHEET_NAME\(\)\)\n    If wsMain Is Nothing Then Exit Sub\n)')
        if pattern_apply.search(content):
            content = pattern_apply.sub(r'\1' + hl_remove_code, content)

    with open(bas_file, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Implemented Option A seamlessly.")

if __name__ == "__main__":
    main()
