import fitz

input_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf"
output_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020_new.pdf"
doc = fitz.open(input_pdf)
font_path = r"C:\Windows\Fonts\arial.ttf"

def process_year(page, pix, old_year, new_year):
    rects = page.search_for(old_year)
    for rect in rects:
        # Sample background color slightly outside the top-left corner
        x = int(max(0, rect.x0 - 2))
        y = int(max(0, rect.y0 - 2))
        try:
            px = pix.pixel(x, y)
            bg_color = (px[0]/255.0, px[1]/255.0, px[2]/255.0)
        except:
            bg_color = (1, 1, 1)

        # Wipe rect
        wipe_rect = rect + (-0.5, -0.5, 0.5, 0.5)
        page.draw_rect(wipe_rect, color=None, fill=bg_color)
        
        # Determine text color (white for table headers, black for graph)
        # Table headers are blue (bg_color ~ (0, 0.38, 0.66)), so text should be white
        # Graph background is white, so text should be black or gray.
        # Original text in table is white, original in graph is dark gray.
        # Let's use a simple luminance check:
        luminance = 0.299*bg_color[0] + 0.587*bg_color[1] + 0.114*bg_color[2]
        if luminance < 0.5:
            text_color = (1, 1, 1) # white
        else:
            text_color = (0.35, 0.35, 0.35) # dark gray to match original chart text

        # Insert new text
        # The font size is roughly the height of the rect
        font_size = rect.height * 0.95
        # The baseline origin is roughly bottom-left
        origin = fitz.Point(rect.x0, rect.y1 - rect.height * 0.15)
        
        page.insert_text(origin, new_year, fontname="ari", fontfile=font_path, fontsize=font_size, color=text_color)

for page in doc:
    pix = page.get_pixmap()
    process_year(page, pix, "2019", "2024")
    process_year(page, pix, "2020", "2025")
    
    if page.number == 12: # the graph page (index 12)
        # Legend patch
        page.draw_rect(fitz.Rect(850, 275, 890, 292), color=None, fill=(1,1,1))
        page.insert_text(fitz.Point(852, 287), "2024", fontname="ari", fontfile=font_path, fontsize=11, color=(0.35, 0.35, 0.35))
        
        page.draw_rect(fitz.Rect(850, 298, 890, 315), color=None, fill=(1,1,1))
        page.insert_text(fitz.Point(852, 310), "2025", fontname="ari", fontfile=font_path, fontsize=11, color=(0.35, 0.35, 0.35))

doc.save(output_pdf)
print("Draw search complete.")
