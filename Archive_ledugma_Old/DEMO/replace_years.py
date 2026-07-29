import fitz

input_pdf = r"C:\ledugma\DEMO\Ad_190x60_PressQuality_v13.pdf"
output_pdf = r"C:\ledugma\DEMO\Ad_190x60_PressQuality_v14.pdf"

doc = fitz.open(input_pdf)
replaced = 0

for xref in range(1, doc.xref_length()):
    if doc.xref_is_stream(xref):
        try:
            stream = doc.xref_stream(xref)
            if b"2019" in stream or b"2020" in stream:
                stream = stream.replace(b"2019", b"2024")
                stream = stream.replace(b"2020", b"2025")
                doc.update_stream(xref, stream)
                replaced += 1
        except Exception as e:
            print(f"Error processing xref {xref}: {e}")

doc.save(output_pdf)
print(f"Done! Replaced years in {replaced} streams. Output saved to {output_pdf}")
