import pdfplumber
import re
import pandas as pd
import os

def get_receipt_data(pdf_path):

    results = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            text = page.extract_text()
            if not text:
                continue

            # --------------------------
            # 1️⃣ Extract เลขที่ (Document No.)
            # --------------------------
            doc_match = re.search(r"เลขที่\s+([A-Z0-9]+)", text)
            doc_no = doc_match.group(1) if doc_match else None

            # --------------------------
            # 2️⃣ Extract วันที่ (Date - Thai format)
            # --------------------------
            date_match = re.search(
                r"วันที่\s+(\d{1,2}\s+[ก-๙]+\s+\d{4})",
                text
            )
            date = date_match.group(1) if date_match else None

            # --------------------------
            # 3️⃣ Extract รวมเงิน (Total)
            # --------------------------
            total_match = re.search(
                r"รวมเงิน\s+([\d,]+)",
                text
            )
            total = total_match.group(1) if total_match else None

            results.append([
                doc_no,
                date,
                total
            ])

    df = pd.DataFrame(results, columns=[
        "เลขที่",
        "วันที่",
        "รวมเงิน"
    ])

    # Save
    path_dir = os.path.dirname(pdf_path)
    file_name = os.path.basename(pdf_path)
    file_name_no_ext = os.path.splitext(file_name)[0]

    output_path = os.path.join(path_dir, f"Extracted_{file_name_no_ext}.xlsx")
    df.to_excel(output_path, index=False)

    print("Done:", output_path)


if __name__ == "__main__":
    # File paths
    
    pdf_path = r"C:\JP_DA\06_K.NICK_PROJECT\01_Henkel\09_4Mall_pdf_summary\Tops(01-01-2025-01-06-2025)_revise1.pdf"
    excel_path = r"C:\JP_DA\06_K.NICK_PROJECT\01_Henkel\09_4Mall_pdf_summary\Tops(01-01-2025-01-06-2025)_revise1.xlsx"

    path_dir = os.path.dirname(pdf_path)
    get_receipt_data(pdf_path)

    # for file in os.listdir(path_dir):
    #     if file.lower().endswith(".pdf"):
    #         print("Processing:", file)
    #         get_receipt_data(os.path.join(path_dir, file))