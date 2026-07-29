with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

print(f"Total lines: {len(text.split('\n'))}")
if "OpenInstallWeb" in text:
    print("OpenInstallWeb exists!")
if "ThisWorkbook.Protect" not in text.split("Sub ExportTotalChart")[1].split("End Sub")[0]:
    print("ExportTotalChart is clean of Protect!")
