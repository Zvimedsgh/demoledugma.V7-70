import win32com.client
excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Add()
ws = wb.Worksheets(1)
shp = ws.Shapes.AddShape(1, 10, 10, 100, 100)
try:
    shp.TextFrame2.MarginRight = 10
    print("TextFrame2.MarginRight exists")
except Exception as e:
    print(f"Error on TextFrame2.MarginRight: {e}")

try:
    shp.TextFrame.MarginRight = 10
    print("TextFrame.MarginRight exists")
except Exception as e:
    print(f"Error on TextFrame.MarginRight: {e}")
wb.Close(False)
excel.Quit()
