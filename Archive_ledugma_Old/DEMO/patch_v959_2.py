import datetime

def create_v959_2():
    with open(r'C:\ledugma\DEMO\V9.59_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    # Add Application.EnableEvents = True to ClearHomePageSelection
    code = code.replace("Public Sub ClearHomePageSelection()\nOn Error Resume Next", 
                        "Public Sub ClearHomePageSelection()\nOn Error Resume Next\nApplication.EnableEvents = True ' Safety reset")

    with open(r'C:\ledugma\DEMO\V9.59_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v959_2()
