import pdfplumber
import pandas as pd
import re
import os
def get_data_villa(pdf_path):

    header_data = {
        "Purchase Order No.": None,
        "Order Date": None,
        "FOR Branch": None,
        "Delivery Date": None
    }

    items = []

    with pdfplumber.open(pdf_path) as pdf:

        for page in pdf.pages:

            text = page.extract_text(layout=True)
            if not text:
                continue

            # ---- Header Extraction ----
            po_match = re.search(
                r"Purchase Order No\./Date\s+([A-Z0-9]+)\s*/\s*([0-9\-]+)",
                text
            )
            if po_match:
                header_data["Purchase Order No."] = po_match.group(1)
                header_data["Order Date"] = po_match.group(2)

            branch_match = re.search(r"FOR Branch\s*:?(\d+)", text)
            if branch_match:
                header_data["FOR Branch"] = branch_match.group(1)

            delivery_match = re.search(r"กำหนดส่งวันที่\s*:?([0-9\-]+)", text)
            if delivery_match:
                header_data["Delivery Date"] = delivery_match.group(1)

            # ---- Item Extraction ----
            lines = text.split("\n")

            for line in lines:

                line = line.strip()

                if not re.match(r"^\d+\s+\d+\s+", line):
                    continue

                code_match = re.search(r"\d{6,}", line)
                if not code_match:
                    continue

                product_code = code_match.group()
                before_code = line.split(product_code)[0]
                product_name = re.sub(r"^\d+\s+\d+\s+", "", before_code).strip()

                after_code = line.split(product_code, 1)[1]
                numbers = re.findall(r"\d[\d,]*\.\d+(?:\s*/\d+)?", after_code)

                if len(numbers) < 7:
                    continue

                unit_pieces = numbers[0]
                unit_per_pack = numbers[1]
                price_per_unit = numbers[2]
                percent_discount = numbers[3]
                discount_per_unit = numbers[4]
                total_price = numbers[5]
                vat = numbers[6]

                items.append({
                    "Purchase Order No.": header_data["Purchase Order No."],
                    "Order Date": header_data["Order Date"],
                    "FOR Branch": header_data["FOR Branch"],
                    "Delivery Date": header_data["Delivery Date"],
                    "Product Name": product_name,
                    "Product Code": product_code,
                    "Unit Pieces": unit_pieces,
                    "Unit per Pack": unit_per_pack,
                    "Price per Unit": price_per_unit,
                    "Percent Discount": percent_discount,
                    "Discount per Unit": discount_per_unit,
                    "Total Price": total_price,
                    "VAT": vat
                })

    df = pd.DataFrame(items)

    print("Rows extracted:", len(df))

    if not df.empty:
        path_dir = os.path.dirname(pdf_path)
        file_name = os.path.basename(pdf_path)
        file_name_no_ext = os.path.splitext(file_name)[0]

        output_path = os.path.join(path_dir, f"Extracted_{file_name_no_ext}.xlsx")
        df.to_excel(output_path, index=False)
        print("Saved to:", output_path)
    else:
        print("No data extracted.")

    return df

if __name__ == "__main__":
    # File paths
    
    pdf_path = r"C:\Users\topde\Downloads\tops\Villa\Villa.pdf"
    get_data_villa(pdf_path)

    # path_dir = os.path.dirname(pdf_path)

    # for i in os.listdir(path_dir):
    #     print(i)
    #     get_data_villa(os.path.join(path_dir, i))
   