import os
import glob
import win32com.client as win32

# Find the file
search_path = r'C:\Users\Zvi\OneDrive*'
found_file = None
for root, dirs, files in os.walk(r'C:\Users\Zvi'):
    if 'OneDrive' in root:
        for f in files:
            if f.lower() == 'manual.docx':
                found_file = os.path.join(root, f)
                break
    if found_file:
        break

if not found_file:
    print("File not found.")
else:
    print(f"Found: {found_file}")
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(found_file)
    
    # Do replacements
    replacements = {
        "> [!NOTE]": "הערה:",
        "> [!TIP]": "טיפ:",
        "> [!WARNING]": "אזהרה חשובה:",
        "> [!SUCCESS]": "הצלחה:"
    }
    
    for find_text, replace_text in replacements.items():
        word.Selection.Find.Execute(FindText=find_text, ReplaceWith=replace_text, Replace=2) # wdReplaceAll=2
    
    # Also clean up any lingering "> " at the start of paragraphs if possible, 
    # but word.Selection.Find.Execute doesn't easily do regex for start of paragraph.
    # It's better to just replace "> " with "" globally, but that might remove valid arrows.
    # Let's just leave the simple ones.
    
    doc.Save()
    doc.Close()
    word.Quit()
    print("Document cleaned successfully!")
