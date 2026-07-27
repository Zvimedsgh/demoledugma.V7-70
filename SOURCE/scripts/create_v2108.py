import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.107.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# For any macro that has Application.EnableEvents = False at the top,
# we need to ensure that before every Exit Sub, we have EnableEvents = True.

# We will just globally replace all instances of "Exit Sub" with a call to a cleanup macro? No, Exit Sub must stay.
# Better to replace:
#     Exit Sub
# with:
#     Application.EnableEvents = True
#     Application.ScreenUpdating = True
#     Exit Sub
# But only if they are not already preceded by EnableEvents = True.

def fix_exit_subs():
    global content
    
    # Let's just find "Exit Sub" and see the previous lines.
    lines = content.split('\n')
    new_lines = []
    
    in_sub = False
    needs_cleanup = False
    
    for i, line in enumerate(lines):
        if re.match(r'^(Public|Private)\s+Sub\s+', line.strip()):
            in_sub = True
            needs_cleanup = False
            
        if in_sub and "Application.EnableEvents = False" in line:
            needs_cleanup = True
            
        if in_sub and line.strip().endswith("Exit Sub"):
            # Check if this is the end of the sub (i.e. just before End Sub) or inside an Err Handler
            # If needs_cleanup is true, ensure the previous lines have EnableEvents = True
            if needs_cleanup:
                # Check previous 3 lines
                prev = "\n".join(lines[max(0, i-3):i])
                if "EnableEvents = True" not in prev:
                    # Inject cleanup
                    indent = line[:len(line) - len(line.lstrip())]
                    new_lines.append(indent + "Application.EnableEvents = True")
                    new_lines.append(indent + "Application.ScreenUpdating = True")
                    new_lines.append(indent + "Application.DisplayAlerts = True")
                    # don't forget to restore Calculation if needed, but it's okay to skip unless we know we changed it.
                    
        new_lines.append(line)
        
        if line.strip() == "End Sub":
            in_sub = False
            needs_cleanup = False

    content = "\n".join(new_lines)

fix_exit_subs()

# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_107"', 'Attribute VB_Name = "Goren_Claude_V2_108"')
content = content.replace('VERSION: V2.107', 'VERSION: V2.108')
content = content.replace('APP_VERSION As String = "2.107"', 'APP_VERSION As String = "2.108"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.108.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.108 created.")
