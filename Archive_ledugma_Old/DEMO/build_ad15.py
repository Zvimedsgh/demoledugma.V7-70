import pypdf
from pypdf import PdfReader, PdfWriter, Transformation, PageObject

text_pdf = PdfReader(r"C:\ledugma\DEMO\text_base_trans_v15.pdf")
slides_pdf = PdfReader(r"C:\ledugma\DEMO\מצגת הנהלה 2020_new.pdf")

# The ad text page
text_page = text_pdf.pages[0]

# Blank page (538.58 x 169.92)
ad_page = PageObject.create_blank_page(width=538.58, height=169.92)

slide13 = slides_pdf.pages[12]
slide14 = slides_pdf.pages[13]
page13 = slides_pdf.pages[12]
page14 = slides_pdf.pages[13]
# Crop the pages to their actual content to prevent white background overlapping
page14.mediabox.lower_left = (10, 0)
page14.mediabox.upper_right = (950, 540)
# width of page14 is now 940

page13.mediabox.lower_left = (60, 0)
page13.mediabox.upper_right = (910, 540)
# --- Left side (Table) ---
# We want to stretch the bottom of the table to match the bottom of the graph visually.
# Graph visual top = 153.08, visual bottom = 33.84.
# Table visual top = 536 * scale_y + ty. Table visual bottom = 197 * scale_y + ty.
# Solving equations gives scale_y = 0.3517, ty = -35.43
scale_x14 = 0.28
scale_y14 = 0.3517
ty14 = -35.43

# --- Right side (Graph) ---
# Top moves down 15 points, bottom moves up 10 points.
# Original height (at scale 0.28) was 540 * 0.28 = 151.2
# New height = 151.2 - 15 - 10 = 126.2
# New vertical scale = 126.2 / 540 = 0.2337
# New bottom = original bottom (18) + 10 = 28
scale_x13 = 0.28
scale_y13 = 0.2337
ty13 = 28

ad_page.merge_transformed_page(page14, Transformation().scale(sx=scale_x14, sy=scale_y14).translate(10, ty14))
ad_page.merge_transformed_page(page13, Transformation().scale(sx=scale_x13, sy=scale_y13).translate(280, ty13))

# Finally, merge the transparent text on top
ad_page.merge_page(text_page)

writer = PdfWriter()
writer.add_page(ad_page)

with open(r"C:\ledugma\DEMO\Ad_190x60_PressQuality_v35.pdf", "wb") as f:
    writer.write(f)

print("Ad v35 created.")
