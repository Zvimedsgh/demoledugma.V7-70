import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Add()
xlmodule = wb.VBProject.VBComponents.Add(1)
xlmodule.CodeModule.AddFromString("""
Sub TestErr()
    On Error GoTo EH
    Err.Raise 999, "Test", "Test error"
    Exit Sub
EH:
    Application.EnableEvents = True
    Dim n As Long
    n = Err.Number
    MsgBox "Err.Number is: " & n
End Sub
""")

try:
    excel.Application.Run("TestErr")
except Exception as e:
    print(f"Failed to run: {e}")

wb.Close(False)
excel.Quit()
