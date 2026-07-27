import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.186.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.187.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Define Hebrew text cleanly for the new install instructions
install_title = "הוראות התקנה מהירות מקישור האינטרנט"
install_welcome = "הגעתם למערכת הדיווח דרך הדפדפן. כדי להפעיל את המערכת בצורה מלאה במחשב שלכם, בצעו:"
install_steps = [
    ("1. הורדה למחשב (חובה):", "כאן בדפדפן, לחצו למעלה על 'קובץ' (File) -> 'שמירה בשם' (Save As) -> 'הורד עותק' (Download a Copy)."),
    ("2. ביטול חסימת אבטחה (Unblock):", "במחשב, לחצו קליק ימני על הקובץ שהורדתם -> 'מאפיינים' (Properties) -> סמנו V בתיבת 'בטל חסימה' (Unblock) ולחצו אישור."),
    ("3. אישור מאקרו:", "פתחו את הקובץ באקסל במחשב (לא בדפדפן!). אם מופיע פס צהוב למעלה, לחצו על 'הפוך תוכן לזמין' (Enable Content)."),
    ("4. הגדרת זמנים:", "במסך הראשי, בחרו שנת בסיס ושנה נוכחית מהתפריטים הנפתחים כדי להתחיל להשתמש במערכת.")
]
button_text = "הבנתי, אל תציג לי יותר גיליון זה"

op_title = "הוראות תפעול למערכת"
op_welcome = "להלן פירוט שלבי העבודה עם 7 כפתורי המערכת הראשיים:"
op_steps = [
    ("כפתור 1 - בדיקת נתונים:", "בודק את הנתונים ומתריע על שגיאות או פערים. ניתן לתקן את השגיאות ישירות, או לשלוח לבכיר לקבלת התיקונים. לאחר מכן יש להקיש את התיקונים מחדש."),
    ("כפתור 2 - החלת תיקונים:", "מחיל את התיקונים שהזנתם ומכין את דוחות הבסיס לעבודה."),
    ("כפתור 3 - הצגת תוצאות:", "מציג את הגיליונות עם כל תוצאות החישובים שבוצעו."),
    ("כפתור 4 - הצגת גיליונות דיווח:", "פותח את גיליונות הדיווח לבדיקה וצפייה נוחה."),
    ("כפתור 5 - הסתרת גיליונות:", "מסתיר את כל גיליונות העזר ומשאיר רק את המסך הראשי נקי ומסודר."),
    ("כפתור 6 - בניית מצגת:", "בונה את המצגת והדוחות הסופיים של המערכת."),
    ("כפתור 7 - הדפסה ושמירה:", "מאפשר להדפיס או לשמור את הדוחות הסופיים כקובץ PDF בצורה מסודרת.")
]

def str_to_chrw(s):
    return " & ".join(f'ChrW({ord(c)})' if ord(c) > 127 else f'"{c}"' for c in s).replace('" & "', "")

def make_vba_assignment(obj, prop, s):
    code = f'{obj}.{prop} = {str_to_chrw(s)}'
    # split long lines
    max_len = 120
    if len(code) <= max_len:
        return code
    parts = code.split(" & ")
    new_line = ""
    current_line = ""
    for i, part in enumerate(parts):
        if i == 0:
            current_line = part
        else:
            if len(current_line) + len(part) + 3 > max_len:
                new_line += current_line + " & _\n        "
                current_line = part
            else:
                current_line += " & " + part
    new_line += current_line
    return new_line

vba_code = f"""Public Sub CreateInstructionSheets()
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    Dim wsInstall As Worksheet
    Dim wsOp As Worksheet
    Dim sInstall As String
    Dim sOp As String
    
    {make_vba_assignment('sInstall', '', 'הוראות_התקנה').replace('sInstall. = ', 'sInstall = ')}
    {make_vba_assignment('sOp', '', 'הוראות_תפעול').replace('sOp. = ', 'sOp = ')}
    
    ' Create or get Install Sheet
    Set wsInstall = Nothing
    Set wsInstall = ThisWorkbook.Worksheets(sInstall)
    If wsInstall Is Nothing Then
        Set wsInstall = ThisWorkbook.Worksheets.Add(After:=wsMain)
        wsInstall.Name = sInstall
    End If
    wsInstall.DisplayRightToLeft = True
    wsInstall.Cells.Clear
    
    ' Create or get Op Sheet
    Set wsOp = Nothing
    Set wsOp = ThisWorkbook.Worksheets(sOp)
    If wsOp Is Nothing Then
        Set wsOp = ThisWorkbook.Worksheets.Add(After:=wsInstall)
        wsOp.Name = sOp
    End If
    wsOp.DisplayRightToLeft = True
    wsOp.Cells.Clear
    
    ' --- Setup Install Sheet ---
    With wsInstall
        {make_vba_assignment('.Range("B2")', 'Value', install_title)}
        .Range("B2").Font.Size = 24
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Color = RGB(0, 51, 102)
        
        {make_vba_assignment('.Range("B4")', 'Value', install_welcome)}
        .Range("B4").Font.Size = 14
        .Range("B4").Font.Bold = True
        
"""

row = 6
for title, desc in install_steps:
    vba_code += f"        {make_vba_assignment(f'.Range(\"B{row}\")', 'Value', title)}\n"
    vba_code += f"        {make_vba_assignment(f'.Range(\"C{row}\")', 'Value', desc)}\n"
    row += 1

vba_code += f"""        
        .Range("B6:B9").Font.Bold = True
        .Range("B6:C9").Font.Size = 12
        .Range("B6:C9").RowHeight = 35
        .Range("B6:C9").VerticalAlignment = xlCenter
        .Columns("B").ColumnWidth = 35
        .Columns("C").ColumnWidth = 100
        .Range("C6:C9").WrapText = True
        
        ' Add Button
        Dim btnInstall As Shape
        For Each btnInstall In .Shapes
            btnInstall.Delete
        Next btnInstall
        Set btnInstall = .Shapes.AddShape(msoShapeRoundedRectangle, .Range("B12").Left, .Range("B12").Top, 300, 40)
        {make_vba_assignment('btnInstall.TextFrame2.TextRange', 'Text', button_text)}
        btnInstall.TextFrame2.TextRange.Font.Size = 14
        btnInstall.TextFrame2.TextRange.Font.Bold = msoTrue
        btnInstall.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        btnInstall.Fill.ForeColor.RGB = RGB(200, 50, 50)
        btnInstall.OnAction = "HideInstallInst"
        
        .Tab.Color = RGB(0, 150, 0)
    End With
    
    ' --- Setup Op Sheet ---
    With wsOp
        {make_vba_assignment('.Range("B2")', 'Value', op_title)}
        .Range("B2").Font.Size = 24
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Color = RGB(0, 51, 102)
        
        {make_vba_assignment('.Range("B4")', 'Value', op_welcome)}
        .Range("B4").Font.Size = 14
        .Range("B4").Font.Bold = True
        
"""

row = 6
for title, desc in op_steps:
    vba_code += f"        {make_vba_assignment(f'.Range(\"B{row}\")', 'Value', title)}\n"
    vba_code += f"        {make_vba_assignment(f'.Range(\"C{row}\")', 'Value', desc)}\n"
    row += 1

vba_code += f"""        
        .Range("B6:B12").Font.Bold = True
        .Range("B6:C12").Font.Size = 12
        .Range("B6:C12").RowHeight = 40
        .Range("B6:C12").VerticalAlignment = xlCenter
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 110
        .Range("C6:C12").WrapText = True
        
        ' Add Button
        Dim btnOp As Shape
        For Each btnOp In .Shapes
            btnOp.Delete
        Next btnOp
        Set btnOp = .Shapes.AddShape(msoShapeRoundedRectangle, .Range("B14").Left, .Range("B14").Top, 300, 40)
        {make_vba_assignment('btnOp.TextFrame2.TextRange', 'Text', button_text)}
        btnOp.TextFrame2.TextRange.Font.Size = 14
        btnOp.TextFrame2.TextRange.Font.Bold = msoTrue
        btnOp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        btnOp.Fill.ForeColor.RGB = RGB(200, 50, 50)
        btnOp.OnAction = "HideOpInst"
        
        .Tab.Color = RGB(0, 150, 0)
    End With
    
    MsgBoxU {str_to_chrw("הגיליונות נוצרו בהצלחה!")}, vbInformation
    On Error GoTo 0
End Sub
"""

# Find the start of CreateInstructionSheets
start_idx = content.find("Public Sub CreateInstructionSheets()")
if start_idx != -1:
    end_idx = content.find("End Sub", start_idx) + 7
    content = content[:start_idx] + vba_code + content[end_idx:]

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_186"', 'Attribute VB_Name = "Goren_Claude_V2_187"')
content = content.replace('VERSION: V2.186', 'VERSION: V2.187')
content = content.replace('APP_VERSION As String = "2.186"', 'APP_VERSION As String = "2.187"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.187 correctly!")
