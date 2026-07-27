import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.112.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Find all subs/functions that contain Application.ScreenUpdating = True
matches = re.finditer(r'(Public Sub|Private Sub|Public Function|Private Function|Sub|Function)\s+([A-Za-z0-9_]+)', content)

for m in matches:
    func_type = m.group(1)
    func_name = m.group(2)
    start_idx = m.end()
    
    # find next sub/function or end of file
    next_match = re.search(r'\n(Public Sub|Private Sub|Public Function|Private Function|Sub|Function)\s+', content[start_idx:])
    if next_match:
        end_idx = start_idx + next_match.start()
    else:
        end_idx = len(content)
        
    func_body = content[start_idx:end_idx]
    
    if "Application.ScreenUpdating = True" in func_body:
        print(f"Contains ScreenUpdating=True: {func_name}")

