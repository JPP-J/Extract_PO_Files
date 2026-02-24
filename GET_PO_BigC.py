
# v1

import pdfplumber
import pandas as pd
import re
import os



def get_data_bigc(pdf_path):

    header_data = {
        "PO Number": None,
        "Order Date": None,
        "Delivery Date": None
    }

    items = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            text = page.extract_text()
            if not text:
                continue

            # -------- HEADER EXTRACTION --------

            # Extract Order Date + Delivery Date
            date_match = re.search(
                r"\d+\s+\w+\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})",
                text
            )
            if date_match:
                header_data["Order Date"] = date_match.group(1)
                header_data["Delivery Date"] = date_match.group(2)

            # Extract PO Number (10 digit at end of header line)
            po_match = re.search(r"\b(\d{10})\b", text)
            if po_match:
                header_data["PO Number"] = po_match.group(1)

            # -------- ITEM TABLE EXTRACTION --------

            lines = text.split("\n")

            for line in lines:

                # Item rows start with running number + barcode
                if re.match(r"^\d+\s+\d{13}\s+", line):

                    # Extract all decimal numbers including comma
                    numbers = re.findall(r"[\d,]+\.\d+", line)
                    numbers = [float(n.replace(",", "")) for n in numbers]

                    if len(numbers) < 6:
                        continue

                    qty = numbers[0]
                    total_qty = numbers[1]
                    price_per_unit = numbers[2]
                    price_per_pack = numbers[3]
                    discount_percent = numbers[4]
                    total_price = numbers[5]

                    # Extract barcode
                    barcode_match = re.search(r"^\d+\s+(\d{13})", line)
                    barcode = barcode_match.group(1) if barcode_match else None

                    # Remove numeric tail to isolate product section
                    line_without_numbers = re.sub(r"[\d,]+\.\d+", "", line)

                    parts = line_without_numbers.split()

                    # After removing running number and barcode
                    # structure:
                    # index 0 = running number
                    # index 1 = barcode
                    # last 3 before decimals = unit, pal, mult

                    running_no = parts[0]
                    barcode = parts[1]

                    unit = parts[-3]
                    pal = parts[-2]
                    mult = parts[-1]

                    # Product name is everything between barcode and unit
                    product_name = " ".join(parts[2:-3])

                    items.append({
                        "PO Number": header_data["PO Number"],
                        "Order Date": header_data["Order Date"],
                        "Delivery Date": header_data["Delivery Date"],
                        "Barcode": barcode,
                        "Product Name": product_name,
                        "Unit": unit,
                        "Pal": pal,
                        "Mult": mult,
                        "Qty": qty,
                        "Total Qty": total_qty,
                        "Price per Unit": price_per_unit,
                        "Price per Pack": price_per_pack,
                        "Discount %": discount_percent,
                        "Total Price (Ex VAT)": total_price
                    })

    df = pd.DataFrame(items)

    print("Rows extracted:", len(df))

    # if not df.empty:
    #     path_dir = os.path.dirname(pdf_path)
    #     file_name = os.path.basename(pdf_path)
    #     file_name_no_ext = os.path.splitext(file_name)[0]

    #     output_path = os.path.join(path_dir, f"Extracted_{file_name_no_ext}.xlsx")
    #     df.to_excel(output_path, index=False)
    #     print("Saved to:", output_path)
    # else:
    #     print("No data extracted.")

    return df

def file_over_file(pdf_path):

    path_dir = os.path.dirname(pdf_path)

    all_dfs = []

    for file in os.listdir(path_dir):

        if file.lower().endswith(".pdf"):   # avoid non-pdf files
            full_path = os.path.join(path_dir, file)

            print("Processing:", file)

            df = get_data_bigc(full_path)

            if not df.empty:
                all_dfs.append(df)

    # Combine everything
    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        print("Total rows:", len(final_df))
    else:
        final_df = pd.DataFrame()
        print("No data found.")

    final_df.head()

    return final_df



if __name__ == "__main__":
    pdf_path = r"Z:\01_DATA\po\Big C\keep\SPS_PO_01122025091031.pdf"
    # df = get_data_bigc(pdf_path)


    # path_dir = os.path.dirname(pdf_path)

    # for i in os.listdir(path_dir):
    #     print(i)
    #     df= get_data_bigc(os.path.join(path_dir, i))

    df = file_over_file(pdf_path)

    

    if not df.empty:
        path_dir = os.path.dirname(pdf_path)
        file_name = os.path.basename(pdf_path)
        file_name_no_ext = os.path.splitext(file_name)[0]

        output_path = os.path.join(path_dir, f"Extracted_{file_name_no_ext}.xlsx")
        df.to_excel(output_path, index=False)
        print("Saved to:", output_path)
    else:
        print("No data extracted.")

