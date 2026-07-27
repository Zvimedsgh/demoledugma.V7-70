import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.199.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.200.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find InitRawColumns
start_idx = -1
end_idx = -1
for i, line in enumerate(lines):
    if "Private Sub InitRawColumns()" in line:
        start_idx = i
    if start_idx != -1 and "End Sub" in line:
        end_idx = i
        break

if start_idx != -1 and end_idx != -1:
    init_raw_lines = lines[start_idx:end_idx+1]
    # Remove from original location
    del lines[start_idx:end_idx+1]
    
    # Find BuildReview in the NEW lines list
    br_idx = -1
    for i, line in enumerate(lines):
        if "Public Sub BuildReview()" in line:
            br_idx = i
            break
            
    if br_idx != -1:
        # Insert before BuildReview
        lines = lines[:br_idx] + ["\n"] + init_raw_lines + ["\n"] + lines[br_idx:]
        
# Replace versions
content = "".join(lines)
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_199"', 'Attribute VB_Name = "Goren_Claude_V2_200"')
content = content.replace('VERSION: V2.199', 'VERSION: V2.200')
content = content.replace('APP_VERSION As String = "2.199"', 'APP_VERSION As String = "2.200"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.200")
