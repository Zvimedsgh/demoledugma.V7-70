$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()
$vbext_ct_StdModule = 1
$module = $wb.VBProject.VBComponents.Add($vbext_ct_StdModule)
$module.CodeModule.AddFromFile("c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas")

# Force compilation
try {
    $excel.Run("BuildPresentation")
} catch {
    Write-Host "Error: $($_.Exception.Message)"
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
