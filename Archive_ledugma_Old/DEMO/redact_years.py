import fitz

input_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf"
output_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020_new.pdf"
doc = fitz.open(input_pdf)

for page in doc:
    # 2019 -> 2024
    rects_19 = page.search_for("2019")
    for r in rects_19:
        page.add_redact_annot(r, text="2024", fontname="helv", fontsize=8, align=1, text_color=(0,0,0), fill=(1,1,1))
    
    # 2020 -> 2025
    rects_20 = page.search_for("2020")
    for r in rects_20:
        page.add_redact_annot(r, text="2025", fontname="helv", fontsize=8, align=1, text_color=(0,0,0), fill=(1,1,1))
    
    page.apply_redactions()

doc.save(output_pdf)
print("Redaction complete.")
