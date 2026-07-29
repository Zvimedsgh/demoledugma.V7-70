$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open("C:\LEVAV PROJECT\SOURCE\Demo_Reports_System_V7.70 (2).xlsm")
    $ws = $wb.Sheets.Item("דף הבית")
    
    $val = $ws.Range("E8").Value()
    Write-Output "Value in E8 is: $val"
    
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
