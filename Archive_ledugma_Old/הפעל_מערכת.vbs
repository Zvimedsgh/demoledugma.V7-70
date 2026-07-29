Dim fso, strPath, strFile, f, objExcel

Set fso = CreateObject("Scripting.FileSystemObject")
strPath = fso.GetParentFolderName(WScript.ScriptFullName)

strFile = ""
' Find the first xlsm or xlsb file in the folder
For Each f In fso.GetFolder(strPath).Files
    If LCase(fso.GetExtensionName(f.Name)) = "xlsm" Or LCase(fso.GetExtensionName(f.Name)) = "xlsb" Then
        strFile = f.Path
        Exit For
    End If
Next

If strFile = "" Then
    MsgBox "לא נמצא קובץ אקסל (xlsm) בתיקייה זו. אנא ודא שקובץ זה נמצא באותה תיקייה עם המערכת.", vbCritical, "שגיאה"
    WScript.Quit
End If

On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "שגיאה בטעינת אקסל.", vbCritical, "שגיאה"
    WScript.Quit
End If
On Error GoTo 0

objExcel.Visible = True
objExcel.Workbooks.Open strFile
