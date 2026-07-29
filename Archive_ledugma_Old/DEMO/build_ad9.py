import pypdf
from pypdf import PdfReader, PdfWriter, Transformation, PageObject

text_pdf = PdfReader(r"C:\ledugma\DEMO\text_base_trans_v9.pdf")
slides_pdf = PdfReader(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")

# The ad text page
text_page = text_pdf.pages[0]

# Blank page (538.58 x 169.92)
ad_page = PageObject.create_blank_page(width=538.58, height=169.92)

slide13 = slides_pdf.pages[12]
slide14 = slides_pdf.pages[13]

scale_x = 250.0 / 960.0
scale_y_graph = 140.6 / 540.0

# Graph on the right
op13 = Transformation().scale(sx=scale_x, sy=scale_y_graph).translate(tx=278.58, ty=25.0)
slide13.add_transformation(op13)
ad_page.merge_page(slide13)

# Table on the left
y_bottom = 190
y_top = 490

slide14.mediabox.lower_left = (0, y_bottom)
slide14.cropbox.lower_left = (0, y_bottom)
slide14.mediabox.upper_right = (960, y_top)
slide14.cropbox.upper_right = (960, y_top)

cropped_height = y_top - y_bottom  # 300

scale_y_table = 140.6 / cropped_height

ty_table = 25.0 - y_bottom * scale_y_table

op14 = Transformation().scale(sx=scale_x, sy=scale_y_table).translate(tx=10.0, ty=ty_table)
slide14.add_transformation(op14)
ad_page.merge_page(slide14)

# Finally, merge the transparent text on top
ad_page.merge_page(text_page)

writer = PdfWriter()
writer.add_page(ad_page)

with open(r"C:\ledugma\DEMO\Ad_190x60_PressQuality_v9.pdf", "wb") as f:
    writer.write(f)

print("Ad v9 created.")
