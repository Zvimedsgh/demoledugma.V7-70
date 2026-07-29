def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    if 'Private Function HebrewToKey' not in content:
        macro = """
Private Function HebrewToKey(ByVal hebText As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As Long
    Dim mapped As String
    
    hebText = Trim$(hebText)
    result = ""
    
    For i = 1 To Len(hebText)
        ch = AscW(Mid$(hebText, i, 1))
        Select Case ch
            Case 1488: mapped = "A"
            Case 1489: mapped = "B"
            Case 1490: mapped = "G"
            Case 1491: mapped = "D"
            Case 1492: mapped = "H"
            Case 1493: mapped = "V"
            Case 1494: mapped = "Z"
            Case 1495: mapped = "CH"
            Case 1496: mapped = "T"
            Case 1497: mapped = "Y"
            Case 1498: mapped = "K"
            Case 1499: mapped = "K"
            Case 1500: mapped = "L"
            Case 1501: mapped = "M"
            Case 1502: mapped = "M"
            Case 1503: mapped = "N"
            Case 1504: mapped = "N"
            Case 1505: mapped = "S"
            Case 1506: mapped = "A"
            Case 1507: mapped = "P"
            Case 1508: mapped = "P"
            Case 1509: mapped = "TZ"
            Case 1510: mapped = "TZ"
            Case 1511: mapped = "K"
            Case 1512: mapped = "R"
            Case 1513: mapped = "SH"
            Case 1514: mapped = "T"
            Case 32:   mapped = "_"
            Case Else
                If (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or (ch >= 48 And ch <= 57) Then
                    mapped = UCase$(Chr$(ch))
                Else
                    mapped = ""
                End If
        End Select
        result = result & mapped
    Next i
    
    Do While InStr(result, "__") > 0
        result = Replace(result, "__", "_")
    Loop
    If Left$(result, 1) = "_" Then result = Mid$(result, 2)
    If Right$(result, 1) = "_" Then result = Left$(result, Len(result) - 1)
    
    HebrewToKey = result
End Function
"""
        with open(bas_file, 'a', encoding='utf-8') as f:
            f.write("\n" + macro)
        print("Fixed Compile Error HebrewToKey!")
    else:
        print("Macro already exists.")

if __name__ == "__main__":
    main()
