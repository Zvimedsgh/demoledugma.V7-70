import fitz

doc = fitz.open(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")
page = doc[12] # Graph
pix = page.get_pixmap()

min_x = 960
max_x = 0
for y in range(pix.height):
    for x in range(pix.width):
        r, g, b = pix.pixel(x, y)[:3]
        if r < 250 or g < 250 or b < 250:
            if x < min_x: min_x = x
            if x > max_x: max_x = x

print(f"Graph content X: {min_x} to {max_x}")
