import pypdf
from pypdf import PdfReader, PdfWriter, Transformation

text_pdf = PdfReader(r"C:\ledugma\DEMO\text_base.pdf")
slides_pdf = PdfReader(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")

# The ad page
ad_page = text_pdf.pages[0]

# Slides
slide7 = slides_pdf.pages[6]
slide8 = slides_pdf.pages[7]

# Target scale and translations
scale_factor = 250.0 / 960.0

# Slide 8 (Left)
op8 = Transformation().scale(sx=scale_factor, sy=scale_factor).translate(tx=12.85, ty=25.0)
slide8.add_transformation(op8)
ad_page.merge_page(slide8)

# Slide 7 (Right)
op7 = Transformation().scale(sx=scale_factor, sy=scale_factor).translate(tx=275.7, ty=25.0)
slide7.add_transformation(op7)
ad_page.merge_page(slide7)

# Save
writer = PdfWriter()
writer.add_page(ad_page)

with open(r"C:\ledugma\DEMO\Ad_190x60_PressQuality_v2.pdf", "wb") as f:
    writer.write(f)

print("Ad created successfully.")
