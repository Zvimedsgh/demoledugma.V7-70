import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.044_20260702_1555.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Add Maps to wsTemp
map_code = """ThisWorkbook.Names.Add "lst_month", wsTemp.Range(wsTemp.Cells(16, 1), wsTemp.Cells(27, 1))

' --- Build Maps for Native Dropdowns ---
wsTemp.Cells(1, 5).Value = tHalf: wsTemp.Cells(1, 6).Value = "lst_half_year"
wsTemp.Cells(2, 5).Value = tQuarter: wsTemp.Cells(2, 6).Value = "lst_quarter"
wsTemp.Cells(3, 5).Value = tMonth: wsTemp.Cells(3, 6).Value = "lst_month"
wsTemp.Cells(4, 5).Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497): wsTemp.Cells(4, 6).Value = "lst_empty"
ThisWorkbook.Names.Add "map_period", wsTemp.Range("E1:F4")

wsTemp.Cells(1, 7).Value = H_COMPANY(): wsTemp.Cells(1, 8).Value = "lst_companies"
wsTemp.Cells(2, 7).Value = H_AGENT(): wsTemp.Cells(2, 8).Value = "lst_agents"
wsTemp.Cells(3, 7).Value = H_TELLER(): wsTemp.Cells(3, 8).Value = "lst_tellers"
wsTemp.Cells(4, 7).Value = H_BRANCH(): wsTemp.Cells(4, 8).Value = "lst_branches"
wsTemp.Cells(5, 7).Value = H_BRANCH() & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494): wsTemp.Cells(5, 8).Value = "lst_main_branches"
ThisWorkbook.Names.Add "map_filter", wsTemp.Range("G1:H5")"""

content = content.replace('ThisWorkbook.Names.Add "lst_month", wsTemp.Range(wsTemp.Cells(16, 1), wsTemp.Cells(27, 1))', map_code)

# 2. Add Native Formulas safely
old_val = """wsMain.Range("G6").MergeArea.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Formula1:="=lst_empty"

wsMain.Range("G7").MergeArea.Validation.Delete"""

new_val = """Dim backupG5 As String
backupG5 = wsMain.Range("G5").Value
wsMain.Range("G5").Value = tHalf
wsMain.Range("G6").MergeArea.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(VLOOKUP($G$5,map_period,2,0))"
wsMain.Range("G5").Value = backupG5

wsMain.Range("G7").MergeArea.Validation.Delete"""

content = content.replace(old_val, new_val)

old_val2 = """wsMain.Range("G9").MergeArea.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Formula1:="=lst_empty"
Err.Clear"""

new_val2 = """Dim backupG8 As String
backupG8 = wsMain.Range("G8").Value
wsMain.Range("G8").Value = H_COMPANY()
wsMain.Range("G9").MergeArea.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(VLOOKUP($G$8,map_filter,2,0))"
wsMain.Range("G8").Value = backupG8
Err.Clear"""

content = content.replace(old_val2, new_val2)

# 3. Disable UpdatePeriodDropdown and UpdateFilterValueDropdown
content = content.replace("Public Sub UpdatePeriodDropdown()\n\n10      Dim wsMain", "Public Sub UpdatePeriodDropdown()\nExit Sub\n\n10      Dim wsMain")
content = content.replace("Public Sub UpdateFilterValueDropdown()\n\n10      Dim wsMain", "Public Sub UpdateFilterValueDropdown()\nExit Sub\n\n10      Dim wsMain")

# 4. Version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_044"', 'Attribute VB_Name = "Goren_Claude_V2_045"')
content = content.replace('Private Const APP_VERSION As String = "2.044"', 'Private Const APP_VERSION As String = "2.045"')
content = content.replace('VERSION: V2.044', 'VERSION: V2.045')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.045 with bulletproof INDIRECT Native validation")
