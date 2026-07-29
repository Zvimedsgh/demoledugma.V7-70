import pypdf
reader = pypdf.PdfReader(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
page = reader.pages[13]

with open(r"C:\ledugma\DEMO\coords_x.txt", "w", encoding="utf-8") as f:
    def visitor_body(text, cm, tm, fontDict, fontSize):
        x = tm[4]
        if text.strip():
            f.write(f"x={x:.1f}, text={text}\n")
    
    page.extract_text(visitor_text=visitor_body)
