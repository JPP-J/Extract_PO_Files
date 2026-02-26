import pdfplumber

pdf_path = r"Z:\01_DATA\po\HomePro-New\PO_0000003074_20260225_091029173970.PDF"
# pdf_path = r"Z:\01_DATA\po\ScanFiles\BigC\DOC010725-004.pdf"

with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[166]
    text = page.extract_text()
    words = page.extract_words()
    print(text)
    # print(words)




# import pytesseract
# from pdf2image import convert_from_path

# pdf_path = r"Z:\01_DATA\po\ScanFiles\BigC\DOC010725-004.pdf"

# pages = convert_from_path(pdf_path)

# for page in pages:
#     text = pytesseract.image_to_string(page)
#     print(text)