import sys
import shutil
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.033.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

shutil.copy(filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("Goren_Claude_V2_033", "Goren_Claude_V2_034")
content = content.replace("V2.033", "V2.034")
content = content.replace('APP_VERSION As String = "2.033"', 'APP_VERSION As String = "2.034"')

# We will use regex to reliably replace ShowHiddenSheets body
pattern = r'(Public Sub ShowHiddenSheets\(\).*?For Each ws In ThisWorkbook\.Worksheets\n)(.*?)(\n\s*Next ws\n\s*On Error GoTo 0)'

def repl(m):
    return m.group(1) + """        If ws.Name <> MATACH_SHEET_NAME() Then
            If IsParameterSheet(ws.Name) Then
                If ws.Visible <> xlSheetHidden Then ws.Visible = xlSheetHidden
            Else
                If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
            End If
        End If""" + m.group(3)

content = re.sub(pattern, repl, content, flags=re.DOTALL)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.034 successfully")
