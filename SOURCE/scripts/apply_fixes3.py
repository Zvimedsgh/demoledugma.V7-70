import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119_temp.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_update = "110     If filterType = H_COMPANY() Then"

new_update = """
    Dim isDemoModeUI As Boolean
    Dim demoParamUI As String
    isDemoModeUI = FORCE_DEMO_MODE
    On Error Resume Next
    demoParamUI = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
    If Not FORCE_DEMO_MODE Then
        If demoParamUI = ChrW(1499) & ChrW(1503) Or demoParamUI = "YES" Then isDemoModeUI = True
        If demoParamUI = ChrW(1500) & ChrW(1488) Or demoParamUI = "NO" Then isDemoModeUI = False
    End If
    On Error GoTo ERR_HANDLER

    If isDemoModeUI Then
        Dim yearValUI As String
        yearValUI = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
        Dim wsDataUI As Worksheet
        On Error Resume Next
        Set wsDataUI = ThisWorkbook.Worksheets("DATA_" & yearValUI)
        On Error GoTo ERR_HANDLER
        
        If Not wsDataUI Is Nothing Then
            Dim dCol As Long
            If filterType = H_COMPANY() Then
                dCol = BASE_COL_COMPANY
            ElseIf filterType = H_TELLER() Then
                dCol = BASE_COL_TELLER
            ElseIf filterType = H_AGENT() Then
                dCol = BASE_COL_AGENTNAME
            ElseIf filterType = H_BRANCH() Then
                dCol = BASE_COL_BRANCHNAME
            ElseIf filterType = H_BRANCH() & ChrW(32) & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494) Then
                dCol = BASE_COL_MAINBRANCH
            Else
                GoTo CLEAN_EXIT
            End If
            
            Dim dLastR As Long
            dLastR = wsDataUI.Cells(wsDataUI.Rows.Count, dCol).End(xlUp).Row
            If dLastR > 1 Then
                Dim dArr As Variant
                dArr = wsDataUI.Range(wsDataUI.Cells(2, dCol), wsDataUI.Cells(dLastR, dCol)).Value2
                
                Dim collUI As Object
                Set collUI = CreateObject("Scripting.Dictionary")
                
                Dim dItem As Variant
                Dim rUI As Long
                If IsArray(dArr) Then
                    For rUI = 1 To UBound(dArr, 1)
                        If Not IsError(dArr(rUI, 1)) Then
                            dItem = Trim$(CStr(dArr(rUI, 1)))
                            If dItem <> "" Then
                                collUI(dItem) = 1
                            End If
                        End If
                    Next rUI
                ElseIf dLastR = 2 Then
                    dItem = Trim$(CStr(dArr))
                    If dItem <> "" Then collUI(dItem) = 1
                End If
                
                If collUI.Count > 0 Then
                    listsName = H_RESHIMOT()
                    Set wsLists = ThisWorkbook.Worksheets(listsName)
                    wsLists.Range("T:T").ClearContents
                    
                    Dim kUI As Variant
                    rUI = 1
                    For Each kUI In collUI.Keys
                        rUI = rUI + 1
                        wsLists.Cells(rUI, 20).Value = kUI
                    Next kUI
                    
                    ThisWorkbook.Names.Add "lst_temp_filter", wsLists.Range(wsLists.Cells(2, 20), wsLists.Cells(rUI, 20))
                    valList = "=lst_temp_filter"
                    GoTo APPLY_VALIDATION
                End If
            End If
        End If
    End If

110     If filterType = H_COMPANY() Then"""

if target_update in content:
    content = content.replace(target_update, new_update)
else:
    print("Could not find target_update.")

target_apply_valid = """' Add validation list to G10
400     wsMain.Range("rngFilterValue").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=valList"""

new_apply_valid = """APPLY_VALIDATION:
' Add validation list to G10
400     wsMain.Range("rngFilterValue").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=valList"""

if target_apply_valid in content:
    content = content.replace(target_apply_valid, new_apply_valid)
else:
    print("Could not find target_apply_valid.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_119"', 'Attribute VB_Name = "Goren_Claude_V2_120"')
content = content.replace('VERSION: V2.119', 'VERSION: V2.120')
content = content.replace('APP_VERSION As String = "2.119"', 'APP_VERSION As String = "2.120"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.120.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.120 created.")
