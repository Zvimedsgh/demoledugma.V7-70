import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.108.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet()" in line:
        in_func = True
    elif in_func and "End Sub" in line:
        break
    elif in_func and "Range(\"A1\").Value =" in line:
        # found where it updates the title. Print everything between the start of A00 and this line.
        # just print all the lines with potentially dangerous code (no error handling)
        pass

# let's just see if there's any dangerous line before 3430
count = 0
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet()" in line:
        in_func = True
    elif in_func and "3430" in line:
        break
    elif in_func:
        # print first few lines of the function to see if there's an On Error GoTo ERR_HANDLER
        if count < 10:
            print(f"[{i+1}] {line.strip()}")
            count += 1
