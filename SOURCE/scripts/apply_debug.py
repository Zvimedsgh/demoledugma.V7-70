import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_err = """ERR_HANDLER:
    ' Just ignore on error
End Sub"""

new_err = """ERR_HANDLER:
    MsgBoxU "Error in UpdateClientList: " & Err.Description & " (Line: " & Erl & ")", vbCritical
End Sub"""

if old_err in content:
    content = content.replace(old_err, new_err)
    print("Injected error reporting to UpdateClientList.")
else:
    print("Could not find ERR_HANDLER block.")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153_debug.bas', 'w', encoding='utf-8') as f:
    f.write(content)

