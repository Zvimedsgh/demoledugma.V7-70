import pypdf
from pypdf import PdfReader, PdfWriter, Transformation

text_pdf = PdfReader(r"C:\ledugma\DEMO\text_base.pdf")
slides_pdf = PdfReader(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")

ad_page = text_pdf.pages[0]

slide13 = slides_pdf.pages[12]
slide14 = slides_pdf.pages[13]

scale_x = 250.0 / 960.0
scale_y_graph = 140.6 / 540.0

# Graph on the right
op13 = Transformation().scale(sx=scale_x, sy=scale_y_graph).translate(tx=278.58, ty=25.0)
slide13.add_transformation(op13)
ad_page.merge_page(slide13)

# Table on the left
# Crop exactly around the table content (y=190 to y=490)
y_bottom = 190
y_top = 490

slide14.mediabox.lower_left = (0, y_bottom)
slide14.cropbox.lower_left = (0, y_bottom)
slide14.mediabox.upper_right = (960, y_top)
slide14.cropbox.upper_right = (960, y_top)

cropped_height = y_top - y_bottom  # 300

scale_y_table = 140.6 / cropped_height

# ty calculation: 
# the crop origin is (0, y_bottom).
# Wait, pypdf transformation applies to original coords!
# So original y_bottom will be scaled to y_bottom * scale_y_table
# We want original y_bottom to map to 25.0
ty_table = 25.0 - y_bottom * scale_y_table

op14 = Transformation().scale(sx=scale_x, sy=scale_y_table).translate(tx=10.0, ty=ty_table)
slide14.add_transformation(op14)
ad_page.merge_page(slide14)

writer = PdfWriter()
writer.add_page(ad_page)

with open(r"C:\ledugma\DEMO\Ad_190x60_PressQuality_v7.pdf", "wb") as f:
    writer.write(f)

print("Ad v7 created.")
