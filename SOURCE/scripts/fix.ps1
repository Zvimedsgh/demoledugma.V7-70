function Replace-Lines {
    param(
        [string]$file,
        [int]$startLine,
        [int]$endLine,
        [string]$replacement
    )
    $lines = Get-Content -Path $file -Raw
    $linesArray = $lines -split "
"
    if ($linesArray.Length -eq 1) { $linesArray = $lines -split "
" }
    
    $newLines = @()
    for ($i = 0; $i -lt $linesArray.Length; $i++) {
        if ($i -lt ($startLine - 1) -or $i -gt ($endLine - 1)) {
            $newLines += $linesArray[$i]
        } elseif ($i -eq ($startLine - 1)) {
            $newLines += $replacement
        }
    }
    Set-Content -Path $file -Value ($newLines -join "
") -Encoding UTF8
}
Replace-Lines -file "C:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.82.bas" -startLine 3212 -endLine 3235 -replacement '
    '' ---- RESTORE MISSING SECTIONS ----
    wsMgmt.Range("A" & (210 + offsetRow) & ":F" & (1000 + offsetRow)).ClearContents
    wsMgmt.Range("A" & (210 + offsetRow) & ":F" & (1000 + offsetRow)).Font.Bold = False
    wsMgmt.Range("A" & (210 + offsetRow) & ":F" & (1000 + offsetRow)).Font.Size = 11
    
    Dim rIgnore As Long, rPeriod As Long, rMsg As Long, rPerm As Long, rClient As Long
    rIgnore = 210 + offsetRow
    rPeriod = 230 + offsetRow
    rMsg = 280 + offsetRow
    rPerm = 300 + offsetRow
    rClient = 320 + offsetRow
    
    '' 1. Ignore Codes
    wsMgmt.Range("A" & rIgnore).Value = ChrW(1511) & ChrW(1493) & ChrW(1491) & ChrW(1497) & " " & ChrW(1492) & ChrW(1514) & ChrW(1506) & ChrW(1500) & ChrW(1502) & ChrW(1493) & ChrW(1514)
    wsMgmt.Range("A" & rIgnore).Font.Bold = True
    wsMgmt.Range("A" & (rIgnore + 1)).Value = "EOD"
    
    '' 2. PeriodLists
    wsMgmt.Range("A" & rPeriod).Value = "PeriodLists"
    wsMgmt.Range("A" & rPeriod).Font.Bold = True
    wsMgmt.Range("A" & (rPeriod + 40)).Value = "EOD"
    
    '' 3. Messages
    wsMgmt.Range("A" & rMsg).Value = ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1493) & ChrW(1514)
    wsMgmt.Range("A" & rMsg).Font.Bold = True
    wsMgmt.Range("A" & (rMsg + 1)).Value = "MSG_DONE"
    wsMgmt.Range("B" & (rMsg + 1)).Value = ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492)
    wsMgmt.Range("A" & (rMsg + 2)).Value = "MSG_ERROR"
    wsMgmt.Range("B" & (rMsg + 2)).Value = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ":"
    wsMgmt.Range("A" & (rMsg + 3)).Value = "EOD"
    
    '' 4. Permissions
    wsMgmt.Range("A" & rPerm).Value = ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1488) & ChrW(1493) & ChrW(1514)
    wsMgmt.Range("A" & rPerm).Font.Bold = True
    wsMgmt.Range("A" & (rPerm + 1)).Value = "zvi"
    wsMgmt.Range("B" & (rPerm + 1)).Value = "fullAccess"
    wsMgmt.Range("A" & (rPerm + 2)).Value = "claude"
    wsMgmt.Range("B" & (rPerm + 2)).Value = "fullAccess"
    wsMgmt.Range("A" & (rPerm + 3)).Value = "EOD"
    
    '' 5. Clients
    wsMgmt.Range("A" & rClient).Value = ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    wsMgmt.Range("A" & rClient).Font.Bold = True
    wsMgmt.Range("A" & (rClient + 1)).Value = "EOD"
'
