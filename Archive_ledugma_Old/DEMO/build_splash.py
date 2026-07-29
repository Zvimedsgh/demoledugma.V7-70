import win32com.client

def build_splash():
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception as e:
        print("Excel is not running.")
        return
    
    wb = excel.ActiveWorkbook
    if wb is None:
        print("No active workbook.")
        return
        
    ws = None
    for s in wb.Worksheets:
        if s.Name == "פתיח":
            ws = s
            break
            
    if ws is None:
        print("Could not find sheet 'פתיח'")
        return
        
    try:
        # Clear existing shapes
        for shp in ws.Shapes:
            shp.Delete()
            
        # Format sheet
        ws.Cells.Interior.Color = 15790320 # Light gray
        excel.ActiveWindow.DisplayGridlines = False
        
        # Add a large button (msoShapeRoundedRectangle = 5)
        shp = ws.Shapes.AddShape(5, 200, 150, 400, 80)
        # Blue: RGB(41, 128, 185) -> 41 + 128*256 + 185*65536 = 12156969
        shp.Fill.ForeColor.RGB = 12156969
        shp.Line.Visible = 0
        
        textframe = shp.TextFrame2
        textframe.TextRange.Text = "להורדת המערכת למחשב - לחץ כאן"
        textframe.TextRange.Font.Size = 24
        textframe.TextRange.Font.Name = "Arial"
        textframe.TextRange.Font.Bold = -1
        textframe.TextRange.Font.Fill.ForeColor.RGB = 16777215 # White
        textframe.VerticalAnchor = 3 # middle
        textframe.TextRange.ParagraphFormat.Alignment = 2 # center
        
        # Add instruction text (msoTextOrientationHorizontal = 1)
        shpText = ws.Shapes.AddTextbox(1, 150, 250, 500, 50)
        shpText.Fill.Visible = 0
        shpText.Line.Visible = 0
        tf2 = shpText.TextFrame2
        tf2.TextRange.Text = "לאחר ההורדה, פתח את הקובץ ולחץ על 'הפעל תוכן' (Enable Content) למעלה כדי להתחיל."
        tf2.TextRange.Font.Size = 16
        tf2.TextRange.Font.Name = "Arial"
        tf2.TextRange.Font.Fill.ForeColor.RGB = 3289650 # Dark gray
        tf2.TextRange.ParagraphFormat.Alignment = 2
        tf2.TextRange.Font.Bold = -1
        
        # We can't easily add hyperlink via win32com without knowing the link, so the user will do it.
        
        ws.Activate()
        print("Splash screen built successfully on 'פתיח'!")
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == '__main__':
    build_splash()
