$xl = New-Object -ComObject Excel.Application
try {
    $wb = $xl.Workbooks.Open("C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70.xlsm")
    $comp = $wb.VBProject.VBComponents.Item("modDemoReports_V9") # or whatever name
    Write-Output "Trust Access Enabled!"
    $wb.Close($false)
} catch {
    Write-Output "Trust Access Disabled: $_"
} finally {
    $xl.Quit()
}
