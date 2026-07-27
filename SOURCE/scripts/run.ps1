$content = Get-Content -Path "c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.993.bas" -Raw

$o = 'Private Sub SetDynamicNamedRange(ws As Worksheet, nameToAssign As String, headerText As String, fallbackRow As Long)
    Dim fnd As Range
    Set fnd = ws.Columns(1).Find(What:=headerText, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not fnd Is Nothing Then
        ThisWorkbook.Names.Add nameToAssign, fnd
    Else
        ThisWorkbook.Names.Add nameToAssign, ws.Range("A" & fallbackRow)
    End If
End Sub'

$n = 'Private Sub SetDynamicNamedRange(ws As Worksheet, nameToAssign As String, headerText As String, fallbackRow As Long)
    Dim fnd As Range
    Dim firstAddress As String
    Set fnd = ws.Columns(1).Find(What:=headerText, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not fnd Is Nothing Then
        firstAddress = fnd.Address
        Do
            If fnd.Font.Bold = True Then
                ThisWorkbook.Names.Add nameToAssign, fnd
                Exit Sub
            End If
            Set fnd = ws.Columns(1).FindNext(fnd)
        Loop While Not fnd Is Nothing And fnd.Address <> firstAddress
    End If
    ThisWorkbook.Names.Add nameToAssign, ws.Range("A" & fallbackRow)
End Sub'

$o = $o -replace '\r\n', "
"
$n = $n -replace '\r\n', "
"
$content = $content -replace '\r\n', "
"

if ($content.Contains($o)) {
    $content = $content.Replace($o, $n)
    
    $o2 = '    SetDynamicNamedRange wsMgmt, "rngSection_FieldMap", ChrW(1502) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1497), 1 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_BranchName", ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501), 57 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_Params", ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512) & ChrW(1497) & ChrW(1501), 173 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_PeriodLists", ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1493) & ChrW(1514), 250 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_Messages", ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1493) & ChrW(1514), 300 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_Permissions", ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1488) & ChrW(1493) & ChrW(1514), 350 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_ReasonCode", ChrW(1505) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1514), 400 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_Clients", ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514), 500 + offsetRow'
    
    $n2 = '    SetDynamicNamedRange wsMgmt, "rngSection_FieldMap", ChrW(1502) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1497), 1 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_BranchName", ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501), 57 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_Params", ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512) & ChrW(1497) & ChrW(1501), 173 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_PeriodLists", ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1493) & ChrW(1514), 250 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_Messages", ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1493) & ChrW(1514), 1000 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_Permissions", ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1488) & ChrW(1493) & ChrW(1514), 1100 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_ReasonCode", ChrW(1505) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1514), 1200 + offsetRow
    SetDynamicNamedRange wsMgmt, "rngSection_Clients", ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514), 1300 + offsetRow'

    $o2 = $o2 -replace '\r\n', "
"
    $n2 = $n2 -replace '\r\n', "
"
    $content = $content.Replace($o2, $n2)

    $content = $content.Replace('Attribute VB_Name = "Goren_Claude1_993"', 'Attribute VB_Name = "Goren_Claude1_994"')
    $content = $content.Replace('VERSION: V1.993', 'VERSION: V1.994')
    $content = $content.Replace('Private Const APP_VERSION As String = "1.993"', 'Private Const APP_VERSION As String = "1.994"')
    $content = $content -replace "
", "
"
    Set-Content -Path "c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.994.bas" -Value $content -Encoding UTF8
    Write-Output "SUCCESS"
} else {
    Write-Output "ORIGINAL STRING NOT FOUND"
}
