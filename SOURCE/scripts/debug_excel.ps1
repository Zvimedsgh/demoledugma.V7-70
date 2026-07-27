try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $excel.ActiveWorkbook
    
    $outFile = "C:\LEVAV PROJECT\SOURCE\names_debug.txt"
    "Active Workbook: " + $wb.Name | Out-File -FilePath $outFile -Encoding UTF8
    
    $names_to_check = @("lst_period_type", "lst_half_year", "lst_quarter", "lst_month")
    
    foreach ($name in $names_to_check) {
        try {
            $n = $wb.Names.Item($name)
            "$name = " + $n.RefersTo | Out-File -FilePath $outFile -Encoding UTF8 -Append
            
            $rn = $excel.Range($name)
            "Values for ${name}:" | Out-File -FilePath $outFile -Encoding UTF8 -Append
            if ($rn.Count -eq 1) {
                "  " + $rn.Value2 | Out-File -FilePath $outFile -Encoding UTF8 -Append
            } else {
                foreach ($cell in $rn) {
                    "  " + $cell.Value2 | Out-File -FilePath $outFile -Encoding UTF8 -Append
                }
            }
        } catch {
            "Error reading ${name}: " + $_.Exception.Message | Out-File -FilePath $outFile -Encoding UTF8 -Append
        }
    }
    Write-Output "Success"
} catch {
    Write-Output "Failed: " + $_.Exception.Message
}
