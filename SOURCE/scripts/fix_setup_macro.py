import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.072.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.073.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_setup = False
setup_lines = []

for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_073"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.073\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.073"\n')
    elif "Public Sub A00_SetupMainSheet()" in line:
        in_setup = True
        new_lines.append(line)
    elif in_setup and "End Sub" in line:
        # Process setup_lines
        setup_content = "".join(setup_lines)
        
        # Strip all existing line numbers in A00_SetupMainSheet
        stripped_lines = []
        for sl in setup_lines:
            sl_stripped = re.sub(r'^\s*\d+\s+', '', sl)
            stripped_lines.append(sl_stripped)
            
        # Add new line numbers to every executable line
        renumbered_lines = []
        line_num = 10
        in_err_handler = False
        for sl in stripped_lines:
            if sl.strip() == "":
                renumbered_lines.append(sl)
            elif sl.strip().startswith("'"):
                renumbered_lines.append(sl)
            elif sl.strip().startswith("Dim "):
                renumbered_lines.append(sl)
            elif sl.strip().startswith("ERR_HANDLER:"):
                in_err_handler = True
                renumbered_lines.append(sl)
            elif in_err_handler:
                renumbered_lines.append(sl) # don't number inside err handler
            else:
                # Add line number
                # Preserve leading whitespace
                match = re.match(r'^(\s*)(.*)', sl)
                indent = match.group(1)
                code = match.group(2)
                renumbered_lines.append(f"{indent}{line_num} {code}\n")
                line_num += 10
        
        # Replace the ERR_HANDLER block
        final_setup = "".join(renumbered_lines)
        err_handler_replacement = """    Dim errNum As Long
    Dim errDesc As String
    Dim errLine As Long
    Dim errSrc As String
    errNum = Err.Number
    errDesc = Err.Description
    errLine = Erl
    errSrc = Err.Source
    Application.EnableEvents = True
    If errNum <> 0 Then
        On Error Resume Next
        MsgBoxU "Error in SetupMainSheet!" & vbCrLf & "Desc: " & errDesc & vbCrLf & "Line: " & errLine & vbCrLf & "Num: " & errNum, vbCritical, "Setup Error"
    End If
"""
        final_setup = re.sub(r'Application\.EnableEvents = True\s*\n\s*If Err\.Number <> 0 Then.*?End If\s*\n', err_handler_replacement, final_setup, flags=re.DOTALL)
        
        new_lines.extend(final_setup.splitlines(True))
        new_lines.append(line)
        in_setup = False
    elif in_setup:
        setup_lines.append(line)
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.073 created with properly numbered SetupMainSheet and robust Err handler.")

