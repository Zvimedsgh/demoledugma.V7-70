import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip_next = 0
for i, line in enumerate(lines):
    if skip_next > 0:
        skip_next -= 1
        continue
    
    if "wsMain.Range(\"G6\").MergeArea.Validation.Add" in line and "Formula1:=\"=lst_empty\"" in lines[i+1]:
        new_lines.append('Dim backupG5 As String\n')
        new_lines.append('backupG5 = wsMain.Range("G5").Value\n')
        new_lines.append('wsMain.Range("G5").Value = tHalf\n')
        new_lines.append('wsMain.Range("G6").MergeArea.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(VLOOKUP($G$5,map_period,2,0))"\n')
        new_lines.append('wsMain.Range("G5").Value = backupG5\n')
        skip_next = 1
    elif "wsMain.Range(\"G9\").MergeArea.Validation.Add" in line and "Formula1:=\"=lst_empty\"" in lines[i+1]:
        new_lines.append('Dim backupG8 As String\n')
        new_lines.append('backupG8 = wsMain.Range("G8").Value\n')
        new_lines.append('wsMain.Range("G8").Value = H_COMPANY()\n')
        new_lines.append('wsMain.Range("G9").MergeArea.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(VLOOKUP($G$8,map_filter,2,0))"\n')
        new_lines.append('wsMain.Range("G8").Value = backupG8\n')
        skip_next = 1
    else:
        new_lines.append(line)

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Injected INDIRECT(VLOOKUP) successfully")
