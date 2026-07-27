import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.071.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.072.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_072"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.072\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.072"\n')
    else:
        new_lines.append(line)

content = "".join(new_lines)

# Find ERR_HANDLER blocks and fix the order
pattern = re.compile(
    r'(ERR_HANDLER:\s*\n)'
    r'(\s*Application\.EnableEvents = True\s*\n)'
    r'(\s*Application\.DisplayAlerts = True\s*\n)?'
    r'(\s*Dim errLine As Long\s*\n)'
    r'(\s*Dim errNum As Long\s*\n)'
    r'(\s*Dim errDesc As String\s*\n)?'
    r'(\s*Dim errSrc As String\s*\n)?'
    r'(\s*errLine = Erl\s*\n)'
    r'(\s*errNum = Err\.Number\s*\n)'
    r'(\s*errDesc = Err\.Description\s*\n)?'
    r'(\s*errSrc = Err\.Source\s*\n)?',
    re.MULTILINE
)

def replacer(match):
    # match.group(1) is ERR_HANDLER:\n
    # group 2 is EnableEvents
    # group 3 is DisplayAlerts
    # groups 4-11 are Dims and assignments
    
    # We want: ERR_HANDLER:\n + Dims + Assignments + EnableEvents + DisplayAlerts
    res = match.group(1)
    res += match.group(4) if match.group(4) else ""
    res += match.group(5) if match.group(5) else ""
    res += match.group(6) if match.group(6) else ""
    res += match.group(7) if match.group(7) else ""
    res += match.group(8) if match.group(8) else ""
    res += match.group(9) if match.group(9) else ""
    res += match.group(10) if match.group(10) else ""
    res += match.group(11) if match.group(11) else ""
    res += match.group(2) if match.group(2) else ""
    res += match.group(3) if match.group(3) else ""
    return res

content_fixed = pattern.sub(replacer, content)

# There might be some blocks with line numbers inside ERR_HANDLER
# Let's write a smarter fixer
new_content_fixed = []
in_err_handler = False
err_vars_captured = False

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content_fixed)

print("V2.072 created with basic regex fix.")

