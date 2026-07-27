import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.048.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip_next = 0
for i, line in enumerate(lines):
    if skip_next > 0:
        skip_next -= 1
        continue
        
    # Revert FORCE_DEMO_MODE to True
    if "Private Const FORCE_DEMO_MODE As Boolean = False" in line:
        new_lines.append(line.replace("False", "True"))
        continue
        
    # Standardize checking logic
    # Look for patterns like `isDemoMode = False` followed by `demoParam = ...`
    if "isDemoMode =" in line and "demoParam =" in lines[i+1] and "GetStringParameter" in lines[i+1]:
        # We replace the next 3 lines with the standardized block
        new_lines.append('isDemoMode = FORCE_DEMO_MODE\n')
        new_lines.append(lines[i+1])
        new_lines.append(lines[i+2].replace('\n', '') + ' Then isDemoMode = True\n')
        new_lines.append('If demoParam = ChrW(1500) & ChrW(1488) Or demoParam = "NO" Then isDemoMode = False\n')
        skip_next = 2
        
        # Sometimes it's written as `If demoParam = ... Then isDemoMode = True`
        # wait, let's just make sure we skip exactly the `If... Then...` line
        if lines[i+2].strip().startswith("If"):
            pass
        continue

    if "isDemoMode =" in line and "demoParam =" in lines[i+2] and "GetStringParameter" in lines[i+2]:
        # Sometimes there's a line in between (e.g., `Dim demoParam As String`)
        new_lines.append('isDemoMode = FORCE_DEMO_MODE\n')
        new_lines.append(lines[i+1])
        new_lines.append(lines[i+2])
        new_lines.append(lines[i+3].replace('\n', '') + ' Then isDemoMode = True\n')
        new_lines.append('If demoParam = ChrW(1500) & ChrW(1488) Or demoParam = "NO" Then isDemoMode = False\n')
        skip_next = 3
        continue

    # Change default parameter insertion in A00_SetupMainSheet
    if "wsParams.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = ChrW(1500) & ChrW(1488)" in line:
        new_lines.append("wsParams.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = ChrW(1499) & ChrW(1503) ' YES\n")
        continue

    # Bump version
    if "Attribute VB_Name =" in line:
        new_lines.append(line.replace("V2_048", "V2_049"))
        continue
    if "Private Const APP_VERSION As String =" in line:
        new_lines.append(line.replace("2.048", "2.049"))
        continue
    if "VERSION: V2.048" in line:
        new_lines.append(line.replace("2.048", "2.049"))
        continue

    new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.049.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.049 with DEMO MODE logic fixed")
