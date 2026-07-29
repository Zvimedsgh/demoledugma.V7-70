import pypdf
from pypdf import PdfReader, PdfWriter, Transformation

text_pdf = PdfReader(r"C:\ledugma\DEMO\text_base.pdf")
slides_pdf = PdfReader(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")

# The ad page
ad_page = text_pdf.pages[0]

# Slides
slide13 = slides_pdf.pages[12]
slide14 = slides_pdf.pages[13]

# We will crop Slide 14 to focus on the table.
# Original size: 960x540
# Let's crop it to remove top title and bottom margin.
# Estimate: table starts at y=150, ends at y=450 (y is from bottom)
# x starts at 50, ends at 910
slide14.cropbox.lower_left = (20, 100)
slide14.cropbox.upper_right = (940, 480)

# Now slide 14 width is 920, height is 380
# Target width for slide 14: let's give it more space.
# Total width available: 538.58
# Slide 13 (graph) width: maybe 200?
# Slide 14 (table) width: maybe 310?
# Margins: 10 + 200 + 10 + 310 + 8.58 = 538.58

# Let's scale Slide 13 (Graph) to 210 width
scale13 = 210.0 / 960.0
op13 = Transformation().scale(sx=scale13, sy=scale13).translate(tx=315.0, ty=30.0)
slide13.add_transformation(op13)
ad_page.merge_page(slide13)

# Let's scale Slide 14 (Table) to 290 width
# Since we cropped it, its internal width is 920.
scale14 = 290.0 / 920.0
# Translating cropped box: The origin is still (0,0) of the original uncropped page!
# So the visible part is from 20 * scale14 to 940 * scale14 in X.
# We want the visible part to start at x=10.
# So tx + 20*scale14 = 10  => tx = 10 - 20*scale14
tx14 = 10.0 - 20 * scale14
# We want the visible part to start at y=30.
ty14 = 30.0 - 100 * scale14

op14 = Transformation().scale(sx=scale14, sy=scale14).translate(tx=tx14, ty=ty14)
slide14.add_transformation(op14)
ad_page.merge_page(slide14)

writer = PdfWriter()
writer.add_page(ad_page)

with open(r"C:\ledugma\DEMO\Ad_190x60_PressQuality_v4.pdf", "wb") as f:
    writer.write(f)

print("Ad v4 created.")
