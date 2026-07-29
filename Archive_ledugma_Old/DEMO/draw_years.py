import fitz

input_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf"
output_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020_new.pdf"
doc = fitz.open(input_pdf)
font_path = r"C:\Windows\Fonts\arial.ttf"

for page in doc:
    pix = page.get_pixmap()
    dict_text = page.get_text("dict")
    
    for block in dict_text.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"]
                color = fitz.sRGB_to_pdf(span["color"]) if isinstance(span["color"], int) else (0,0,0)
                
                if "2019" in text or "2020" in text:
                    rect = fitz.Rect(span["bbox"])
                    
                    # Get local background color
                    # Sample slightly inside the bounding box, or just outside the top-left
                    # Actually, top-left slightly outside is safest to avoid hitting anti-aliased text pixels
                    x = int(max(0, rect.x0 - 2))
                    y = int(max(0, rect.y0 - 2))
                    try:
                        px = pix.pixel(x, y)
                        bg_color = (px[0]/255.0, px[1]/255.0, px[2]/255.0)
                    except:
                        bg_color = (1,1,1)
                        
                    # draw background to wipe out text
                    # We slightly expand the rect to cover the text fully without leaving ghost outlines
                    wipe_rect = rect + (-0.5, -0.5, 0.5, 0.5)
                    page.draw_rect(wipe_rect, color=None, fill=bg_color)
                    
                    # draw new text
                    new_text = text.replace("2019", "2024").replace("2020", "2025")
                    page.insert_text(fitz.Point(span["origin"]), new_text, fontname="ari", fontfile=font_path, fontsize=span["size"], color=color)

doc.save(output_pdf)
print("Dynamic draw complete.")
