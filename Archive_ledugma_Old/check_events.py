with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

if 'Workbook_NewSheet' in text:
    print("Workbook_NewSheet event exists in .bas!")
else:
    print("No Workbook_NewSheet in .bas.")

import os
if os.path.exists(r'c:\LEVAV PROJECT\SOURCE\ThisWorkbook.cls'):
    with open(r'c:\LEVAV PROJECT\SOURCE\ThisWorkbook.cls', 'r', encoding='utf-8') as f:
        cls_text = f.read()
    if 'Workbook_NewSheet' in cls_text:
        print("Workbook_NewSheet event exists in ThisWorkbook.cls!")
    else:
        print("No Workbook_NewSheet in ThisWorkbook.cls.")
