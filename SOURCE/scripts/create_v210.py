import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.209.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.210.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

def inject_replace_safe(func_name, content):
    start_idx = content.find(f"Private Function {func_name}() As String")
    if start_idx == -1: return content
    end_idx = content.find("End Function", start_idx)
    if end_idx == -1: return content
    
    block = content[start_idx:end_idx+12]
    new_block = block.replace("End Function", f"{func_name} = Replace({func_name}, \":\\\\\\\\\", \":\\\\\")\n    {func_name} = Replace({func_name}, \":\\\\\\\\\", \":\\\\\")\nEnd Function")
    return content[:start_idx] + new_block + content[end_idx+12:]

content = inject_replace_safe("SOURCE_FOLDER", content)
content = inject_replace_safe("REPORTS_FOLDER", content)
content = inject_replace_safe("BACKUP_FOLDER", content)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_209"', 'Attribute VB_Name = "Goren_Claude_V2_210"')
content = content.replace('VERSION: V2.209', 'VERSION: V2.210')
content = content.replace('APP_VERSION As String = "2.209"', 'APP_VERSION As String = "2.210"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.210")
