import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.172.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.173.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    lines = f.readlines()

start_idx = -1
end_idx = -1
for i, line in enumerate(lines):
    if "Public Sub FixUIButtons()" in line:
        start_idx = i
    if "End Sub" in line and start_idx != -1 and i > start_idx + 10:
        end_idx = i + 1
        break

if start_idx != -1 and end_idx != -1:
    del lines[start_idx:end_idx]
    
for i in range(len(lines)):
    lines[i] = lines[i].replace('Attribute VB_Name = "Goren_Claude_V2_172"', 'Attribute VB_Name = "Goren_Claude_V2_173"')
    lines[i] = lines[i].replace('VERSION: V2.172', 'VERSION: V2.173')
    lines[i] = lines[i].replace('APP_VERSION As String = "2.172"', 'APP_VERSION As String = "2.173"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Created V2.173 without FixUIButtons!")
