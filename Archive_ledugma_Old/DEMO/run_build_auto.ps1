Start-Process python -ArgumentList "C:\ledugma\DEMO\click_msgbox.py" -WindowStyle Hidden

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    Write-Output "Opening patched xlsm..."
    $wb = $xl.Workbooks.Open("C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70_patched.xlsm")
    Write-Output "Running BuildPresentation..."
    $xl.Run("Demo_Reports_Syatem_V7.70_patched.xlsm!BuildPresentation")
    Write-Output "Presentation Built!"
    $wb.Close($false)
} catch {
    Write-Output "Error: $_"
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    
    # Kill the clicker
    Stop-Process -Name "python" -Force -ErrorAction SilentlyContinue
}
