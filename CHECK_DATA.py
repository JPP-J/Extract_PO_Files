import pdfplumber

pdf_path = r"C:\JP_DA\06_K.NICK_PROJECT\01_Henkel\09_4Mall_pdf_summary\Tops(01-01-2025-01-06-2025)_revise1.pdf"
# pdf_path = r"Z:\01_DATA\po\ScanFiles\BigC\DOC010725-004.pdf"

with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[0]
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