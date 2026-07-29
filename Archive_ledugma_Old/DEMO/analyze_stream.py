import fitz
import re

input_pdf = r"C:\ledugma\DEMO\מצגת הנהלה 2020.pdf"
doc = fitz.open(input_pdf)

for xref in range(1, doc.xref_length()):
    if doc.xref_is_stream(xref):
        try:
            stream = doc.xref_stream(xref)
            # Find any stream that has "2" and "0" and "1" and "9" near each other
            if b"2" in stream and b"0" in stream and b"1" in stream and b"9" in stream:
                # Let's print a small snippet to see the format
                # We'll use regex to find combinations like (20) ... (19)
                matches = re.findall(b'\([^\)]*2[^\)]*\).*?\([^\)]*0[^\)]*\)', stream)
                if matches:
                    print(f"xref {xref}: found complex 2019 pattern: {matches[0]}")
                matches2 = re.findall(b'\<[0-9A-Fa-f\s]*\>', stream)
                # print any hex strings that might be '2019'
        except Exception as e:
            pass
