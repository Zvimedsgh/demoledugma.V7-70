import pypdf
from pypdf import PdfReader, PdfWriter, Transformation

text_pdf = PdfReader(r"C:\ledugma\DEMO\text_base.pdf")
slides_pdf = PdfReader(r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf")

# The ad page
ad_page = text_pdf.pages[0]

# Slides
slide13 = slides_pdf.pages[12]
slide14 = slides_pdf.pages[13]

# Target scale and translations
scale_factor = 250.0 / 960.0

# Slide 14 (Left)
op14 = Transformation().scale(sx=scale_factor, sy=scale_factor).translate(tx=12.85, ty=25.0)
slide14.add_transformation(op14)
ad_page.merge_page(slide14)

# Slide 13 (Right)
op13 = Transformation().scale(sx=scale_factor, sy=scale_factor).translate(tx=275.7, ty=25.0)
slide13.add_transformation(op13)
ad_page.merge_page(slide13)

# Save
writer = PdfWriter()
writer.add_page(ad_page)

with open(r"C:\ledugma\DEMO\Ad_190x60_PressQuality_v3.pdf", "wb") as f:
    writer.write(f)

print("Ad created successfully.")
