import os

src_path = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.11.bas'
out_path = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.12.bas'

with open(src_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version header
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_11"', 'Attribute VB_Name = "Goren_Claude_V3_12"')
content = content.replace("' VERSION: V3.11", "' VERSION: V3.12\n' CHANGES IN V3.12:\n'   - BUGFIX: Moved CheckUserPermissions to above Exit Sub in A00_SetupMainSheet so returning to main refreshes permissions.")

# Move CheckUserPermissions
old_text = """4960 Exit Sub
ERR_HANDLER:"""

new_text = """4950 CheckUserPermissions
4960 Exit Sub
ERR_HANDLER:"""

content = content.replace(old_text, new_text)

# And remove it from the very end of ERR_HANDLER to avoid running it twice if there's an error
old_err_text = """    If errNum <> 0 Then
        On Error Resume Next
        MsgBoxU "Error in V3.10! (If you don't see this, you are running old code!)" & vbCrLf & "Desc: " & errDesc & vbCrLf & "Line: " & errLine & vbCrLf & "Num: " & errNum, vbCritical, "Setup Error"
    End If
    CheckUserPermissions

End Sub"""

new_err_text = """    If errNum <> 0 Then
        On Error Resume Next
        MsgBoxU "Error in Setup! " & vbCrLf & "Desc: " & errDesc & vbCrLf & "Line: " & errLine & vbCrLf & "Num: " & errNum, vbCritical, "Setup Error"
    End If
    CheckUserPermissions ' Keep here just in case error happened before it

End Sub"""

content = content.replace(old_err_text, new_err_text)

with open(out_path, 'w', encoding='utf-8') as f:
    f.write(content)

print('V3.12 generated successfully')
