import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_str = """ThisWorkbook.Names.Add "lst_period_type", wsTemp.Range(wsTemp.Cells(3, 1), wsTemp.Cells(6, 1))"""
new_str = """ThisWorkbook.Names.Add "lst_empty", wsTemp.Range(wsTemp.Cells(50, 1), wsTemp.Cells(50, 1))
ThisWorkbook.Names.Add "lst_period_type", wsTemp.Range(wsTemp.Cells(3, 1), wsTemp.Cells(6, 1))"""

if old_str in content:
    content = content.replace(old_str, new_str)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Added lst_empty")
else:
    print("Could not find string")
