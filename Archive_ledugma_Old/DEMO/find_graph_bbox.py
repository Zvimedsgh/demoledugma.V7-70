import fitz

doc = fitz.open(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
page = doc[12] # Graph
pix = page.get_pixmap()

min_y = 540
max_y = 0
for y in range(pix.height):
    for x in range(pix.width):
        r, g, b = pix.pixel(x, y)[:3]
        if r < 250 or g < 250 or b < 250:
            if y < min_y: min_y = y
            if y > max_y: max_y = y

print(f"Graph content Y: {min_y} to {max_y}")
