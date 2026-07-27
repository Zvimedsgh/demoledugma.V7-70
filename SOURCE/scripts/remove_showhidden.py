import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.155.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_call = "1650: Call ShowHiddenSheets"
# Wait, it might not have line numbers in the actual file!
# In the actual file it's just "Call ShowHiddenSheets" around line 1650.
# Let's read lines and find it.
lines = content.split('\n')
out_lines = []
for i, line in enumerate(lines):
    if "Call ShowHiddenSheets" in line and "UpdateClientList" in lines[i-1]:
        out_lines.append(line.replace("Call ShowHiddenSheets", "' REMOVED: Call ShowHiddenSheets (was unhiding all base/data sheets incorrectly)"))
    else:
        out_lines.append(line)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write('\n'.join(out_lines))

print("Removed Call ShowHiddenSheets")
