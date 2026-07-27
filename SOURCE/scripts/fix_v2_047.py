import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.046_20260702_1615.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip_next = 0
for i, line in enumerate(lines):
    if skip_next > 0:
        skip_next -= 1
        continue
        
    # 1. Revert INDIRECT(VLOOKUP) to =lst_empty in A00_SetupMainSheet
    if "wsMain.Range(\"G6\").MergeArea.Validation.Add Type:=xlValidateList" in line and "INDIRECT" in line:
        new_lines.append('wsMain.Range("G6").MergeArea.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_empty"\n')
        continue
    if "backupG5 = wsMain.Range(\"G5\").Value" in line or "wsMain.Range(\"G5\").Value = tHalf" in line or "wsMain.Range(\"G5\").Value = backupG5" in line:
        continue
    if "Dim backupG5 As String" in line:
        continue

    if "wsMain.Range(\"G9\").MergeArea.Validation.Add Type:=xlValidateList" in line and "INDIRECT" in line:
        new_lines.append('wsMain.Range("G9").MergeArea.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_empty"\n')
        continue
    if "backupG8 = wsMain.Range(\"G8\").Value" in line or "wsMain.Range(\"G8\").Value = H_COMPANY()" in line or "wsMain.Range(\"G8\").Value = backupG8" in line:
        continue
    if "Dim backupG8 As String" in line:
        continue

    # 2. Remove Exit Sub from UpdatePeriodDropdown
    if "Public Sub UpdatePeriodDropdown()" in line:
        new_lines.append(line)
        # Check if next line is Exit Sub
        if "Exit Sub" in lines[i+1]:
            skip_next = 1
        continue
        
    # 3. Remove Exit Sub from UpdateFilterValueDropdown
    if "Public Sub UpdateFilterValueDropdown()" in line:
        new_lines.append(line)
        if "Exit Sub" in lines[i+1]:
            skip_next = 1
        continue
        
    # 4. Fix "riv'oni" to "riv'on"
    if "ElseIf InStr(1, periodType, ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497)" in line:
        new_lines.append('110     ElseIf InStr(1, periodType, ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503), vbTextCompare) > 0 Then\n')
        continue

    new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.047.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.047")
