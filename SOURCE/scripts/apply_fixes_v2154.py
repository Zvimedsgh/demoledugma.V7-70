import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target = "1440                If Abs(premVal) > threshold And Not dictCorrections.Exists(CStr(r)) Then GoTo NextSrcRow"

if target in content:
    content = content.replace(target, "1440                ' Removed buggy premVal filter")
    print("Fixed BuildBaseSheet by removing the buggy filter.")
else:
    print("Could not find target string.")

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

