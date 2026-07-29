import re

def fix_code():
    with open(r'C:\ledugma\DEMO\V9.49_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()
    
    # 1. MsgBox error lines
    code = code.replace('" G10", vbExclamation', '" G9", vbExclamation')
    
    # 2. btnSearchTop
    code = code.replace('btnSearchTop = wsMain.Range("F13").Top + wsMain.Range("F13").Height + 5',
                        'btnSearchTop = wsMain.Range("F11").Top + wsMain.Range("F11").Height + 15')
    
    # 3. Setup Filter Type
    code = code.replace('wsMain.Range("G10").Value =', 'wsMain.Range("G9").Value =')
    code = code.replace('wsMain.Range("G10").Validation', 'wsMain.Range("G9").Validation')
    code = code.replace('IsEmpty(wsMain.Range("G10").Value) Or wsMain.Range("G10").Value = ""',
                        'IsEmpty(wsMain.Range("G9").Value) Or wsMain.Range("G9").Value = ""')
                        
    # 4. Setup Filter Value
    code = code.replace('wsMain.Range("G11").Validation', 'wsMain.Range("G10").Validation')
    
    # 5. Client Name label & value
    code = code.replace('wsMain.Range("F13")', 'wsMain.Range("F11")')
    code = code.replace('wsMain.Range("G13")', 'wsMain.Range("G11")')
    code = code.replace('IsEmpty(wsMain.Range("G13").Value) Or wsMain.Range("G13").Value = ""',
                        'IsEmpty(wsMain.Range("G11").Value) Or wsMain.Range("G11").Value = ""')
    
    # 6. Named Ranges
    code = code.replace('RefersTo:="=\'" & wsMain.Name & "\'!$G$10"', 'RefersTo:="=\'" & wsMain.Name & "\'!$G$9"')
    code = code.replace('RefersTo:="=\'" & wsMain.Name & "\'!$G$11"', 'RefersTo:="=\'" & wsMain.Name & "\'!$G$10"')
    code = code.replace('RefersTo:="=\'" & wsMain.Name & "\'!$G$13"', 'RefersTo:="=\'" & wsMain.Name & "\'!$G$11"')
    
    # 7. Borders and Interior Colors
    code = code.replace('wsMain.Range("F4:G13").Borders', 'wsMain.Range("F4:G11").Borders')
    code = code.replace('wsMain.Range("G4:G13").Interior.Color', 'wsMain.Range("G4:G11").Interior.Color')
    code = code.replace('wsMain.Range("F4:F13").Interior.Color', 'wsMain.Range("F4:F11").Interior.Color')
    
    # 8. Conditional Formatting
    code = code.replace('wsMain.Range("G11").FormatConditions.Delete', 'wsMain.Range("G10").FormatConditions.Delete')
    code = code.replace('wsMain.Range("G11").FormatConditions.Add Type:=xlExpression, Formula1:="=AND($G$11="""", $G$10<>"""")"',
                        'wsMain.Range("G10").FormatConditions.Add Type:=xlExpression, Formula1:="=AND($G$10="""", $G$9<>"""")"')
    code = code.replace('wsMain.Range("G11").FormatConditions(wsMain.Range("G11").FormatConditions.count).Interior.Color',
                        'wsMain.Range("G10").FormatConditions(wsMain.Range("G10").FormatConditions.count).Interior.Color')
    
    # 9. TestInit updates
    code = code.replace('\' G10: filter value = bechar/i\n    wsMain.Range("G11").Value', '\' G10: filter value = bechar/i\n    wsMain.Range("G10").Value')
    
    # Update comments just to be neat
    code = code.replace('Filter area setup (F9:G11)', 'Filter area setup (F9:G10)')
    code = code.replace('Client name filter (F12 label, G12 Data Validation dropdown)', 'Client name filter (F11 label, G11 Data Validation dropdown)')
    
    with open(r'C:\ledugma\DEMO\V9.50_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)
    
    print("Created V9.50_CopyPaste.txt")

if __name__ == "__main__":
    fix_code()
