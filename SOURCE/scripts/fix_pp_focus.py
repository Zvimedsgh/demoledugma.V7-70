import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.156.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = """If resMsg = vbYes Then
        ppApp.Visible = True
        ppApp.Activate
        With ppPres.SlideShowSettings"""

new_code = """If resMsg = vbYes Then
        Application.WindowState = xlMinimized
        ppApp.Visible = True
        On Error Resume Next
        ppApp.WindowState = 2 ' Min
        ppApp.WindowState = 3 ' Max
        On Error GoTo ERR_HANDLER
        ppApp.Activate
        With ppPres.SlideShowSettings"""

content = content.replace(old_code, new_code)

old_code_no = """ElseIf resMsg = vbNo Then
        ppApp.Visible = True
        ppApp.Activate
    Else"""

new_code_no = """ElseIf resMsg = vbNo Then
        Application.WindowState = xlMinimized
        ppApp.Visible = True
        On Error Resume Next
        ppApp.WindowState = 2
        ppApp.WindowState = 3
        On Error GoTo ERR_HANDLER
        ppApp.Activate
    Else"""

content = content.replace(old_code_no, new_code_no)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("PowerPoint focus fix applied.")
