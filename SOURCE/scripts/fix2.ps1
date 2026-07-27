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

Replace-Lines -file "C:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.84.bas" -startLine 3215 -endLine 3252 -replacement '
    '' ---- NO MORE WIPING ----
    '' Removed the destructive wipe so user settings are preserved.
'

Replace-Lines -file "C:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.84.bas" -startLine 3778 -endLine 3794 -replacement '
        '' Reverted back to using named ranges directly since wsMgmt is no longer VeryHidden!
'

Replace-Lines -file "C:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.84.bas" -startLine 3795 -endLine 3795 -replacement '200     wsMain.Range("rngPeriodValue").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & listName'

