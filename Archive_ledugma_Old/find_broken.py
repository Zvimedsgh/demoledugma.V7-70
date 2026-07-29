with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

import re

for i, line in enumerate(lines):
    l = line.strip()
    if not l or l.startswith("'") or l.startswith("Rem ") or l.startswith("#"):
        continue
    
    # Strip line numbers
    l = re.sub(r'^\d+\s+', '', l)
    
    # If the line starts with a lowercase letter and is multiple words, and has no assignment
    if re.match(r'^[a-z][a-zA-Z0-9_]+\s+[a-zA-Z0-9_]+', l) and '=' not in l and '.' not in l and '(' not in l and '&' not in l:
        print(f"Line {i+1}: {l}")
        
    # Or if it starts with an uppercase word but isn't a VBA keyword
    keywords = ['If', 'ElseIf', 'Else', 'End', 'For', 'Next', 'Do', 'Loop', 'While', 'Wend', 'Select', 'Case', 'Exit', 'Dim', 'ReDim', 'Set', 'Call', 'On', 'Resume', 'Goto', 'Sub', 'Function', 'Property', 'Public', 'Private', 'Global', 'Const', 'Type', 'Enum', 'Attribute', 'Option', 'Debug', 'MsgBox', 'MsgBoxU', 'InputBox', 'Err', 'Kill', 'FileCopy', 'MkDir', 'RmDir', 'ChDir', 'ChDrive', 'Open', 'Close', 'Print', 'Write', 'Line', 'Input', 'Get', 'Put', 'Seek', 'Lock', 'Unlock', 'Name', 'Reset', 'Stop', 'SendKeys', 'Beep', 'AppActivate', 'Shell', 'DoEvents', 'Load', 'Unload', 'SaveSetting', 'GetSetting', 'DeleteSetting', 'Application', 'ThisWorkbook', 'ActiveWorkbook', 'ActiveSheet', 'Worksheets', 'Sheets', 'Range', 'Cells', 'Rows', 'Columns', 'With', 'Array', 'String', 'Integer', 'Long', 'Double', 'Boolean', 'Variant', 'Object', 'Date', 'Time', 'True', 'False', 'Nothing', 'Empty', 'Null', 'Optional', 'ByVal', 'ByRef', 'ParamArray', 'Static', 'New', 'As', 'Like', 'Is', 'Not', 'And', 'Or', 'Xor', 'Eqv', 'Imp', 'Mod']
    
    first_word = l.split()[0] if l.split() else ''
    if first_word and first_word[0].isupper() and first_word not in keywords and '=' not in l and '.' not in l and '(' not in l and '&' not in l and not l.endswith(':'):
        print(f"Line {i+1}: {l}")
