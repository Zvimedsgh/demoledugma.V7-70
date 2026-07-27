import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_str = """    If dictMainBranches.Count > 0 Then
        arrKeys = dictMainBranches.keys
        For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 5).Value = arrKeys(kk): Next kk
    End If"""

new_str = """    If dictMainBranches.Count > 0 Then
        arrKeys = dictMainBranches.keys
        For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 5).Value = arrKeys(kk): Next kk
    End If
    
    ' Define Named Ranges for Filter Dropdowns
    On Error Resume Next
    ThisWorkbook.Names("lst_companies").Delete
    ThisWorkbook.Names("lst_tellers").Delete
    ThisWorkbook.Names("lst_agents").Delete
    ThisWorkbook.Names("lst_branches").Delete
    ThisWorkbook.Names("lst_main_branches").Delete
    On Error GoTo ERR_HANDLER
    
    If dictCompanies.Count > 0 Then ThisWorkbook.Names.Add "lst_companies", wsLists.Range(wsLists.Cells(2, 1), wsLists.Cells(dictCompanies.Count + 1, 1))
    If dictTellers.Count > 0 Then ThisWorkbook.Names.Add "lst_tellers", wsLists.Range(wsLists.Cells(2, 2), wsLists.Cells(dictTellers.Count + 1, 2))
    If dictAgents.Count > 0 Then ThisWorkbook.Names.Add "lst_agents", wsLists.Range(wsLists.Cells(2, 3), wsLists.Cells(dictAgents.Count + 1, 3))
    If dictBranches.Count > 0 Then ThisWorkbook.Names.Add "lst_branches", wsLists.Range(wsLists.Cells(2, 4), wsLists.Cells(dictBranches.Count + 1, 4))
    If dictMainBranches.Count > 0 Then ThisWorkbook.Names.Add "lst_main_branches", wsLists.Range(wsLists.Cells(2, 5), wsLists.Cells(dictMainBranches.Count + 1, 5))"""

if old_str in content:
    content = content.replace(old_str, new_str)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Added named ranges")
else:
    print("Could not find string to replace")
