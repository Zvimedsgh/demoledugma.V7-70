$excel = [Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
if ($excel) {
    $wb = $excel.ActiveWorkbook
    $ws = $wb.Sheets.Item("דף הבית")
    
    # Check if we can unprotect
    try {
        $ws.Unprotect("Z961814r")
    } catch {}

    # Add a round button (msoShapeOval = 9 or msoShapeRoundedRectangle = 5) near MATACH. 
    # MATACH is usually around A1 or somewhere top left.
    # Let's put it at Top=10, Left=10, Width=80, Height=30
    # Wait, the sheet is RTL, so Left=10 is on the physical right. 
    # To put it on the physical left, we need Left to be high, or just near A1/B2.
    # Actually, let's put it at cell B2 or C2.
    $targetCell = $ws.Range("B2")
    $shape = $ws.Shapes.AddShape(5, $targetCell.Left, $targetCell.Top, 120, 30) # 5 = msoShapeRoundedRectangle
    
    $shape.Name = "BtnManual"
    $shape.TextFrame2.TextRange.Text = "הוראות תפעול"
    
    # Format shape
    $shape.Fill.ForeColor.RGB = 16750899 # Light Blue or similar
    $shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0 # Black
    $shape.TextFrame2.TextRange.Font.Name = "Rubik"
    $shape.TextFrame2.TextRange.Font.Size = 11
    $shape.TextFrame2.TextRange.Font.Bold = $true
    
    # Center text
    $shape.TextFrame2.VerticalAnchor = 3 # msoAnchorMiddle
    $shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 # msoAlignCenter
    
    # Assign macro
    $shape.OnAction = "OpenManual"
    
    # Re-protect
    try {
        $ws.Protect("Z961814r")
    } catch {}
    
    Write-Output "Button added successfully!"
} else {
    Write-Output "Excel is not open."
}
