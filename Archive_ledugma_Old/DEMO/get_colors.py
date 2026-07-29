import fitz

doc = fitz.open(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
page = doc[13] # page 14 of the presentation (index 13)
pix = page.get_pixmap()

# Find the 2019 hits
rects = page.search_for("2019")
for r in rects:
    # Get color at the top-left corner of the rect (slightly outside the text)
    x = int(r.x0 - 2)
    y = int(r.y0 - 2)
    if x >= 0 and y >= 0 and x < pix.width and y < pix.height:
        pixel = pix.pixel(x, y)
        print(f"Hit at {r.x0:.1f}, {r.y0:.1f}: BG color = {pixel}")
