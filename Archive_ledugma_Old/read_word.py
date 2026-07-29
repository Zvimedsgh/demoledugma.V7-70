import os
import win32com.client as win32

# Find the file
found_file = None
for root, dirs, files in os.walk(r'C:\Users\Zvi'):
    if 'OneDrive' in root:
        for f in files:
            if f.lower() == 'manual.docx':
                found_file = os.path.join(root, f)
                break
    if found_file:
        break

if found_file:
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(found_file)
    text = doc.Content.Text
    print(text[:2000].encode('utf-8', errors='replace').decode('utf-8'))
    doc.Close(False)
    word.Quit()
