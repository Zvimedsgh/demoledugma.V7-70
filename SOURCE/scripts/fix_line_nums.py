import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.072.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.074.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_setup = False
setup_lines = []

for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_074"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.074\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.074"\n')
    elif "Public Sub A00_SetupMainSheet()" in line:
        in_setup = True
        new_lines.append(line)
    elif in_setup and "End Sub" in line:
        # Process setup_lines
        
        # Strip all existing line numbers in A00_SetupMainSheet
        stripped_lines = []
        for sl in setup_lines:
            sl_stripped = re.sub(r'^\s*\d+\s+', '', sl)
            stripped_lines.append(sl_stripped)
            
        # Add new line numbers to every executable line
        renumbered_lines = []
        line_num = 10
        in_err_handler = False
        is_continued = False
        
        for sl in stripped_lines:
            # We must remember if the PREVIOUS line was continued.
            was_continued = is_continued
            
            # Check if THIS line ends with continuation character `_`
            # Note: it might end with `_ \n` or `_\r\n`
            sl_clean = sl.strip()
            is_continued = sl_clean.endswith("_")
            
            if sl_clean == "":
                renumbered_lines.append(sl)
            elif sl_clean.startswith("'"):
                renumbered_lines.append(sl)
            elif sl_clean.startswith("Dim "):
                renumbered_lines.append(sl)
            elif sl_clean.startswith("ERR_HANDLER:"):
                in_err_handler = True
                renumbered_lines.append(sl)
            elif in_err_handler:
                renumbered_lines.append(sl) # don't number inside err handler
            elif was_continued:
                # If the previous line had `_`, this line is part of the same logical line
                # So we DO NOT add a line number here!
                renumbered_lines.append(sl)
            else:
                # Add line number
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

print("V2.074 created with fixed line continuations.")

