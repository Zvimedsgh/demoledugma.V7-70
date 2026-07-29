try {
    $excel = [Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $excel.ActiveWorkbook
    if ($wb -ne $null) {
        $ws = $wb.Sheets.Item("דף הבית")
        $output = "Workbook: " + $wb.Name + "`n"
        $output += "Checking shapes on Home Page...`n"
        foreach ($shape in $ws.Shapes) {
            $output += "Shape Name: " + $shape.Name + " | Visible: " + $shape.Visible + " | Type: " + $shape.Type + "`n"
            if ($shape.Type -eq 1 -or $shape.Type -eq 17) { # AutoShape or TextBox
                try {
                    $text = $shape.TextFrame2.TextRange.Text
                    if ($text.Length -gt 0) {
                        $output += "  Text snippet: " + $text.Substring(0, [math]::Min($text.Length, 30)) + "...`n"
                    }
                } catch {}
            }
        }
        Write-Output $output
    } else {
        Write-Output "No active workbook."
    }
} catch {
    Write-Output "Excel is not running or COM is blocked."
}
