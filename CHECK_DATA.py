import pdfplumber

pdf_path = r"C:\Users\topde\Downloads\tops\TM\The mall1.pdf"

with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[3]
    text = page.extract_text()
    words = page.extract_words()
    print(text)
    # print(words)