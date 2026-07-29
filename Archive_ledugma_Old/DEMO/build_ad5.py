import pypdf
from pypdf import PdfReader, PdfWriter, Transformation

text_pdf = PdfReader(r"C:\ledugma\DEMO\text_base.pdf")
slides_pdf = PdfReader(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")

# The ad page
ad_page = text_pdf.pages[0]

# Slides
slide13 = slides_pdf.pages[12]
slide14 = slides_pdf.pages[13]

# Equal width for both sides
scale_x = 250.0 / 960.0
scale_y_graph = 250.0 / 960.0

# Graph (Slide 13) on the right
# tx = 538.58 - 10 - 250 = 278.58
op13 = Transformation().scale(sx=scale_x, sy=scale_y_graph).translate(tx=278.58, ty=25.0)
slide13.add_transformation(op13)
ad_page.merge_page(slide13)

# Table (Slide 14) on the left
# Crop the table to remove blank space at bottom and top
# Table data roughly from y=150 to y=480 (height = 330)
slide14.cropbox.lower_left = (0, 150)
slide14.cropbox.upper_right = (960, 480)

# We want the table to be 250 wide (same as graph)
# And we want it stretched downwards to match the graph's height (140.6)
# Graph scaled height = 540 * scale_y_graph = 140.625
# Table cropped height = 330
scale_y_table = 140.625 / 330.0

# Translate so it aligns properly
# Crop origin is still (0,0). The visible part is y=150 to 480.
# We want the visible part to start at ty=25.0
# ty_table + 150 * scale_y_table = 25.0  => ty_table = 25.0 - 150 * scale_y_table
ty_table = 25.0 - 150 * scale_y_table

op14 = Transformation().scale(sx=scale_x, sy=scale_y_table).translate(tx=10.0, ty=ty_table)
slide14.add_transformation(op14)
ad_page.merge_page(slide14)

writer = PdfWriter()
writer.add_page(ad_page)

with open(r"C:\ledugma\DEMO\Ad_190x60_PressQuality_v5.pdf", "wb") as f:
    writer.write(f)

print("Ad v5 created.")
