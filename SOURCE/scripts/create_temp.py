import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.196.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.196_temp.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix ApplyCorrections
old_app = 'Public Sub ApplyCorrectionsAndBuildReports()\n\n10      Dim wsMain As Worksheet'
new_app = 'Public Sub ApplyCorrectionsAndBuildReports()\n    InitRawColumns\n\n10      Dim wsMain As Worksheet'
content = content.replace(old_app, new_app)

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created temp")
