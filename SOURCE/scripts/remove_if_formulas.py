import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.044.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Formula1:=\"=IF($G$5=" in line:
        new_lines.append(line.replace("Formula1:=\"=IF($G$5=\"\"\" & tHalf & \"\"\",lst_half_year,IF($G$5=\"\"\" & tQuarter & \"\"\",lst_quarter,IF($G$5=\"\"\" & tMonth & \"\"\",lst_month,lst_empty)))\"", "Formula1:=\"=lst_empty\""))
    elif "Formula1:=\"=IF($G$8=" in line:
        new_lines.append(line.replace("Formula1:=\"=IF($G$8=\"\"\" & tBranch & \"\"\",lst_branches,IF($G$8=\"\"\" & tAgent & \"\"\",lst_agents,IF($G$8=\"\"\" & tTeller & \"\"\",lst_tellers,IF($G$8=\"\"\" & tComp & \"\"\",lst_companies,IF($G$8=\"\"\" & tMainBranch & \"\"\",lst_main_branches,lst_empty)))))\"", "Formula1:=\"=lst_empty\""))
    else:
        new_lines.append(line)

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Removed complex IF formulas from A00_SetupMainSheet")
