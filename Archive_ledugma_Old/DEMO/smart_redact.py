import fitz

input_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf"
output_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020_new.pdf"
doc = fitz.open(input_pdf)

for page in doc:
    font_path = r"C:\Windows\Fonts\arial.ttf"
    page.insert_font(fontname="ari", fontfile=font_path)
    
    dict_text = page.get_text("dict")
    for block in dict_text.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"]
                color = fitz.sRGB_to_pdf(span["color"]) if isinstance(span["color"], int) else (0,0,0)
                if "2019" in text:
                    rect = fitz.Rect(span["bbox"])
                    page.add_redact_annot(rect, text=text.replace("2019", "2024"), fontname="ari", fontsize=span["size"], text_color=color, fill=(1,1,1))
                if "2020" in text:
                    rect = fitz.Rect(span["bbox"])
                    page.add_redact_annot(rect, text=text.replace("2020", "2025"), fontname="ari", fontsize=span["size"], text_color=color, fill=(1,1,1))
    
    page.apply_redactions()

doc.save(output_pdf)
print("Smart redaction complete.")
