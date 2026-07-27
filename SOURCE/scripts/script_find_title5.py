import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

for i, line in enumerate(content.splitlines()):
    if "Range(\"A2\").Value" in line or "wsMain.Range(\"A1" in line:
        print(f"Line {i+1}: {line}")
