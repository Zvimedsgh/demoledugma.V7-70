import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.190.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.192.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Rebuild CreateInstructionSheets exactly as in V2.191 but with fixed quote escaping
start_idx = content.find("Public Sub CreateInstructionSheets()")
end_idx = content.find("End Sub", start_idx) + 7

def str_to_chrw(s):
    res = []
    for c in s:
        if ord(c) > 127:
            res.append(f'ChrW({ord(c)})')
        elif c == '"':
            res.append('ChrW(34)')
        else:
            res.append(f'"{c}"')
    return " & ".join(res).replace('" & "', "")

def make_vba_assignment(obj, prop, s):
    code = f'{obj}.{prop} = {str_to_chrw(s)}'
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
    ("כפתור 2 - יישום ודו\"חות:", "מחיל את התיקונים שהזנתם ומכין את דוחות הבסיס לעבודה."),
    ("כפתור 3 - הצג גיליונות דיווח:", "מציג את גיליונות הדיווח לבדיקה וצפייה נוחה (למקרה שהוסתרו)."),
    ("כפתור 4 - ייצור מצגת:", "בונה את המצגת והדוחות הסופיים של המערכת."),
    ("כפתור 5 - שמירת דוחות:", "שומר את הגיליונות בדיסק של המחשב."),
    ("כפתור 6 - צפייה בדוחות שמורים:", "פותח לצפייה את הדוחות ששמרנו קודם במחשב."),
    ("כפתור 7 - לקוחות חדשים:", "הפקת דוח של לקוחות שלא היו פעילים בתקופת הבסיס וכן פעילים בשנה הנוכחית.")
]

vba_code = f"""Public Sub CreateInstructionSheets()
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    ' Reset the "Don't Show" flags so they appear for new users!
    wsMain.Range("AA1:AA2").ClearContents
    
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
        .Columns("B").ColumnWidth = 35
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
    
    wsInstall.Activate
    MsgBoxU {str_to_chrw("הגיליונות נוצרו בהצלחה! שמור את הקובץ כעת כשהוא על גיליון ההתקנה.")}, vbInformation
    On Error GoTo 0
End Sub"""

content = content[:start_idx] + vba_code + content[end_idx:]

# 2. Update HideWorkSheets
old_hide1 = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
    If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then 
        hideIt = False
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    End If
End If
If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
    If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA2").Value) <> "YES" Then 
        hideIt = False
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    End If
End If"""

new_hide1 = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
    If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA2").Value) <> "YES" Then 
        hideIt = False
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    End If
End If"""

content = content.replace(old_hide1, new_hide1)

# 3. Update CheckUserPermissions
old_hide2 = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
    If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA1").Value) <> "YES" Then 
        hideIt2 = False
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    End If
End If
If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
    If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA2").Value) <> "YES" Then 
        hideIt2 = False
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    End If
End If"""

new_hide2 = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
    If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA2").Value) <> "YES" Then 
        hideIt2 = False
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    End If
End If"""

content = content.replace(old_hide2, new_hide2)

# 4. Update the jump logic at the end of CheckUserPermissions
old_jump = """    ' But if instruction sheets are enabled, jump to them instead
    Dim sInstall2 As String
    sInstall2 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
    If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA1").Value) <> "YES" Then
        On Error Resume Next
        ThisWorkbook.Worksheets(sInstall2).Activate
        On Error GoTo 0
    End If"""

new_jump = """    ' But if operation instruction sheet is enabled, jump to it instead
    Dim sOp2 As String
    sOp2 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
    If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA2").Value) <> "YES" Then
        On Error Resume Next
        ThisWorkbook.Worksheets(sOp2).Activate
        On Error GoTo 0
    End If"""

content = content.replace(old_jump, new_jump)


content = content.replace('Attribute VB_Name = "Goren_Claude_V2_190"', 'Attribute VB_Name = "Goren_Claude_V2_192"')
content = content.replace('VERSION: V2.190', 'VERSION: V2.192')
content = content.replace('APP_VERSION As String = "2.190"', 'APP_VERSION As String = "2.192"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.192 correctly!")
