import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.047.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "UpdatePeriodDropdown" in line or "110     ElseIf" in line or "lst_empty" in line:
        if "Sub" in line or "ChrW(1512)" in line or "AlertStyle:=xlValidAlertStop" in line:
            print(f"[{i}] {lines[i].strip()}")
