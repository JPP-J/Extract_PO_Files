# v1

import pdfplumber
import pandas as pd
import re
import os

def to_float(value):
    return float(value.replace(",", ""))

def get_data_themall(pdf_path):

    header_data = {
        "PO Number": None,
        "Date Approve": None,
        "Date Delivery": None
    }

    items = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            text = page.extract_text()
            if not text:
                continue

            # -----------------------------------
            # Clean CID encoding garbage
            # -----------------------------------
            text = re.sub(r"\(cid:\d+\)", "", text)

            # -----------------------------------
            # Header extraction
            # -----------------------------------
            po_match = re.search(r"ใบส่งั ซื้อเลขที่\s*(\d+)", text)
            if po_match:
                header_data["PO Number"] = po_match.group(1)

            date_match = re.search(r"Approve\s+([0-9]{2}\.[0-9]{2}\.[0-9]{4})", text)
            if date_match:
                header_data["Date Approve"] = date_match.group(1)

            delivery_match = re.search(r"ส่งสิน.*?:\s*([0-9\.]+)", text)
            if delivery_match:
                header_data["Date Delivery"] = delivery_match.group(1)

            # -----------------------------------
            # Item extraction
            # -----------------------------------
            lines = text.split("\n")

            i = 0
            while i < len(lines):

                line = lines[i].strip()

                # Detect numeric product row
                if re.match(r"^\d{6,12}\s+\d+\s+\d+\s+\d+", line):

                    parts = re.split(r"\s+", line)

                    if len(parts) < 12:
                        i += 1
                        continue

                    product_code = parts[0]
                    pcs_per_pack = int(parts[1])
                    pack = int(parts[2])
                    pack_unit = int(parts[3])

                    price_exclude_discount = to_float(parts[4])
                    discount_pack = to_float(parts[5])
                    discount1 = to_float(parts[6])
                    discount2 = to_float(parts[7])
                    discount3 = to_float(parts[8])

                    qty = int(parts[9])
                    net_price = to_float(parts[10])
                    net_cost = to_float(parts[11])

                    tag = None
                    if len(parts) > 12 and re.match(r"^[A-Za-z]$", parts[12]):
                        tag = parts[12]

                    # -------------------------
                    # Product name + barcode
                    # -------------------------
                    product_name = ""
                    barcode = None

                    j = i + 1

                    while j < len(lines):

                        next_line = lines[j].strip()

                        # Stop if next numeric product row begins
                        if re.match(r"^\d{6,12}\s", next_line):
                            break

                        # Stop at summary section
                        if "รวมเงิน" in next_line:
                            break

                        barcode_match = re.search(r"\d{10,13}", next_line)

                        if barcode_match:
                            barcode = barcode_match.group()
                            clean_name = next_line.replace(barcode, "").strip()
                            product_name += " " + clean_name
                        else:
                            product_name += " " + next_line

                        j += 1

                    items.append({
                        "PO Number": header_data["PO Number"],
                        "Date Approve": header_data["Date Approve"],
                        "Date Delivery": header_data["Date Delivery"],
                        "Product Code": product_code,
                        "Product Name": product_name.strip(),
                        "Barcode": barcode,
                        "Pcs per Pack": pcs_per_pack,
                        "Pack": pack,
                        "Pack Unit": pack_unit,
                        "Price Exclude Discount": price_exclude_discount,
                        "Discount per Pack": discount_pack,
                        "Discount 1": discount1,
                        "Discount 2": discount2,
                        "Discount 3": discount3,
                        "Qty": qty,
                        "Net Price (Exclude VAT)": net_price,
                        "Net Cost After Discount": net_cost,
                        "Tag": tag
                    })

                    i = j
                    continue

                i += 1

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
    pdf_path = r"Z:\01_DATA\po\The mall\The mall6.pdf"
    get_data_themall(pdf_path)


    # path_dir = os.path.dirname(pdf_path)

    # for i in os.listdir(path_dir):
    #     print(i)
    #     get_data_themall(os.path.join(path_dir, i))
   
