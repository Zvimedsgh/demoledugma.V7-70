import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.091.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.092.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if "' LOGGER FOR DEBUGGING HANGS" in line:
        skip = True
        continue
    if skip and "End Sub" in line:
        skip = False
        continue
    if skip:
        continue
        
    # Replace version strings
    line = line.replace('Attribute VB_Name = "Goren_Claude_V2_091"', 'Attribute VB_Name = "Goren_Claude_V2_092"')
    line = line.replace("' VERSION: V2.091", "' VERSION: V2.092")
    line = line.replace('Private Const APP_VERSION As String = "2.091"', 'Private Const APP_VERSION As String = "2.092"')
    
    # Remove the empty line that might be left over from the skipped block
    if not skip and "' -------------------------------------------------------------------------" in line and i > 0 and "' LOGGER FOR DEBUGGING HANGS" in lines[i+1]:
        skip = True
        continue
        
    new_lines.append(line)

logger_func = """
' -------------------------------------------------------------------------
' LOGGER FOR DEBUGGING HANGS
' -------------------------------------------------------------------------
Private Sub LogDebug(ByVal msg As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open "c:\\LEVAV PROJECT\\SOURCE\\debug_log.txt" For Append As #ff
    Print #ff, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & msg
    Close #ff
    On Error GoTo 0
End Sub
"""
new_lines.append(logger_func)

with open(outpath, 'w', encoding='utf-8') as f:
    for line in new_lines:
        f.write(line)

print("V2.092 created with LogDebug at the end.")
