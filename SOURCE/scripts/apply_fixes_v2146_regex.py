import sys, re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.145.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern = re.compile(r"If isDemoMode Then\s*wsMain\.Range\(\"A1\"\)\.Font\.Color = RGB\(200, 0, 0\) ' Red for Demo\s*Else\s*wsMain\.Range\(\"A1\"\)\.Font\.Color = RGB\(0, 0, 0\) ' Black for Real\s*End If")

if pattern.search(content):
    content = pattern.sub(r"wsMain.Range(\"A1\").Font.Color = RGB(200, 0, 0) ' Always Red", content)
    print("Fixed title color logic via regex.")
else:
    print("Could not find title color logic via regex.")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.146.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.146 updated.")
