import fitz

doc = fitz.open(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
output_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020_new.pdf"

# We are going to erase the extra text for Migdal and Kash
# For Migdal: " (המגן) חברה לביטוח"
# For Kash: " בינלאומיים"

for page in doc:
    pix = page.get_pixmap()
    
    # Process Migdal extra text
    rects = page.search_for("המגן")
    for r in rects:
        # Wipe the exact word with a small margin
        wipe_rect = fitz.Rect(r.x0 - 2, r.y0, r.x1 + 2, r.y1)
        x = int(max(0, wipe_rect.x0 - 2))
        y = int(max(0, wipe_rect.y0 - 2))
        try:
            px = pix.pixel(x, y)
            bg_color = (px[0]/255.0, px[1]/255.0, px[2]/255.0)
        except:
            bg_color = (1, 1, 1)
        page.draw_rect(wipe_rect, color=None, fill=bg_color)
        
    rects = page.search_for("לביטוח")
    for r in rects:
        wipe_rect = fitz.Rect(r.x0 - 2, r.y0, r.x1 + 2, r.y1)
        x = int(max(0, wipe_rect.x0 - 2))
        y = int(max(0, wipe_rect.y0 - 2))
        try:
            px = pix.pixel(x, y)
            bg_color = (px[0]/255.0, px[1]/255.0, px[2]/255.0)
        except:
            bg_color = (1, 1, 1)
        page.draw_rect(wipe_rect, color=None, fill=bg_color)
        
    rects = page.search_for("חברה")
    for r in rects:
        wipe_rect = fitz.Rect(r.x0 - 2, r.y0, r.x1 + 2, r.y1)
        x = int(max(0, wipe_rect.x0 - 2))
        y = int(max(0, wipe_rect.y0 - 2))
        try:
            px = pix.pixel(x, y)
            bg_color = (px[0]/255.0, px[1]/255.0, px[2]/255.0)
        except:
            bg_color = (1, 1, 1)
        page.draw_rect(wipe_rect, color=None, fill=bg_color)

    # Process Kash
    rects = page.search_for("בינלאומיים")
    for r in rects:
        wipe_rect = fitz.Rect(r.x0 - 2, r.y0, r.x1 + 2, r.y1)
        x = int(max(0, wipe_rect.x0 - 2))
        y = int(max(0, wipe_rect.y0 - 2))
        try:
            px = pix.pixel(x, y)
            bg_color = (px[0]/255.0, px[1]/255.0, px[2]/255.0)
        except:
            bg_color = (1, 1, 1)
        page.draw_rect(wipe_rect, color=None, fill=bg_color)

doc.save(output_pdf)
print("Text wiped.")
