import pypdf
reader = pypdf.PdfReader(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
page = reader.pages[13]

with open(r"C:\ledugma\DEMO\coords.txt", "w", encoding="utf-8") as f:
    def visitor_body(text, cm, tm, fontDict, fontSize):
        y = tm[5]
        if text.strip():
            f.write(f"y={y:.1f}, text={text}\n")
    
    page.extract_text(visitor_text=visitor_body)
