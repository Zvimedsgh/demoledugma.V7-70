import fitz

input_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf"
output_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020_new.pdf"
doc = fitz.open(input_pdf)
font_path = r"C:\Windows\Fonts\arial.ttf"

for page in doc:
    dict_text = page.get_text("dict")
    replacements = []
    
    for block in dict_text.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"]
                color = fitz.sRGB_to_pdf(span["color"]) if isinstance(span["color"], int) else (0,0,0)
                
                if "2019" in text:
                    rect = fitz.Rect(span["bbox"])
                    replacements.append({
                        "rect": rect,
                        "origin": fitz.Point(span["origin"]),
                        "text": text.replace("2019", "2024"),
                        "size": span["size"],
                        "color": color
                    })
                elif "2020" in text:
                    rect = fitz.Rect(span["bbox"])
                    replacements.append({
                        "rect": rect,
                        "origin": fitz.Point(span["origin"]),
                        "text": text.replace("2020", "2025"),
                        "size": span["size"],
                        "color": color
                    })
    
    # 1. Add redactions to remove old text (no fill)
    for rep in replacements:
        page.add_redact_annot(rep["rect"], cross_out=False) # Default has no fill/text, so it just deletes the content
        
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    
    # 2. Insert new text using our custom font
    page.insert_font(fontname="ari", fontfile=font_path)
    for rep in replacements:
        page.insert_text(rep["origin"], rep["text"], fontname="ari", fontsize=rep["size"], color=rep["color"])

doc.save(output_pdf)
print("Hybrid redaction complete.")
