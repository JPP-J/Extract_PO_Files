import pdfplumber

pdf_path = r"C:\Users\topde\Downloads\tops\Villa\Villa.pdf"

with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[0]
    text = page.extract_text()
    print(text)