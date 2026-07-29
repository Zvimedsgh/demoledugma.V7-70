import win32com.client

def main():
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        wb = xl.ActiveWorkbook
        
        target_ws = None
        target_shp = None
        for ws in wb.Worksheets:
            try:
                shp = ws.Shapes("shpInstallMsg")
                target_ws = ws
                target_shp = shp
                break
            except Exception:
                pass
                
        if not target_shp:
            print("Shape 'shpInstallMsg' not found in any sheet.")
            return

        text = target_shp.TextFrame2.TextRange.Text
        width = target_shp.Width
        height = target_shp.Height
        
        print("WIDTH:", width)
        print("HEIGHT:", height)
        print("TEXT:", repr(text))
        
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    main()
