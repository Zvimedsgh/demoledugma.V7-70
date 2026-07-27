import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.133.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

found_typo = False
found_align = False
found_wa = False

for i, line in enumerate(lines):
    if "ChrW(1513) & \":\"" in line:
        found_typo = True
    if "wsSearch.Range(\"E5:E8\").HorizontalAlignment = -4152" in line:
        found_align = True
    if "\" Whatsapp  \" & ChrW(1500) & \" - 054-6677396\"" in line:
        found_wa = True

print(f"Typo fixed: {found_typo}")
print(f"Alignment fixed: {found_align}")
print(f"Whatsapp fixed: {found_wa}")

