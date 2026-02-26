import pdfplumber
import pandas as pd
import re
import os



# def get_data_homepro(pdf_path):

#     header_data = {
#         "PO Number": None,
#         "Order Date": None,
#         "Delivery Date": None
#     }

#     items = []

#     with pdfplumber.open(pdf_path) as pdf:

#         for page_num, page in enumerate(pdf.pages):

#             text = page.extract_text()
#             print("Reading page:", page_num + 1)

#             if not text:
#                 continue

#             # --------------------------------------------------
#             # HEADER EXTRACTION (only set once)
#             # --------------------------------------------------

#             if not header_data["PO Number"]:
#                 po_match = re.search(r"PO\s*#\s*:\s*(\d+)", text)
#                 if po_match:
#                     header_data["PO Number"] = po_match.group(1)

#             if not header_data["Order Date"]:
#                 order_match = re.search(r"วันที่สั่ง\s*:\s*(\d{2}/\d{2}/\d{4})", text)
#                 if not order_match:
#                     order_match = re.search(r"(\d{2}/\d{2}/\d{4})", text)
#                 if order_match:
#                     header_data["Order Date"] = order_match.group(1)

#             if not header_data["Delivery Date"]:
#                 delivery_match = re.search(r"กำหนดส่ง\s*:\s*(\d{2}/\d{2}/\d{4})", text)
#                 if delivery_match:
#                     header_data["Delivery Date"] = delivery_match.group(1)

#             # --------------------------------------------------
#             # ITEM EXTRACTION
#             # --------------------------------------------------

#             lines = text.split("\n")

#             for line in lines:

#                 line = line.strip()

#                 # Must start with item number
#                 if not re.match(r"^\d+\s+", line):
#                     continue

#                 # Must contain price structure
#                 if "/" not in line:
#                     continue

#                 # Extract numeric values (price & total)
#                 numbers = re.findall(r"[\d,]+\.\d+", line)

#                 if len(numbers) < 2:
#                     continue

#                 price_per_unit = float(numbers[0].replace(",", ""))
#                 total_price = float(numbers[-1].replace(",", ""))

#                 # Extract quantity (number before EA)
#                 qty_match = re.search(r"\s(\d+)\s+EA", line)
#                 if not qty_match:
#                     continue

#                 qty = float(qty_match.group(1))

#                 # Remove price section to isolate product info
#                 left_part = line.split(str(numbers[0]))[0].strip()
#                 parts = left_part.split()

#                 # Remove item number
#                 parts = parts[1:]

#                 # Remove optional tags like [M] or [S]
#                 parts = [p for p in parts if not re.match(r"\[.*\]", p)]

#                 # Find last numeric as product code
#                 product_code = None
#                 for p in reversed(parts):
#                     if p.isdigit():
#                         product_code = p
#                         break

#                 if not product_code:
#                     continue

#                 index = parts.index(product_code)
#                 product_name = " ".join(parts[:index])

#                 items.append({
#                     "PO Number": header_data["PO Number"],
#                     "Order Date": header_data["Order Date"],
#                     "Delivery Date": header_data["Delivery Date"],
#                     "Product Name": product_name,
#                     "Product Code": product_code,
#                     "Price per Unit": price_per_unit,
#                     "Qty": qty,
#                     "Total Price (Ex VAT)": total_price
#                 })

#     df = pd.DataFrame(items)

#     print("Rows extracted:", len(df))

#     return df


# ====================
def get_data_homepro(pdf_path):

    items = []
    current_po = None
    order_date = None
    delivery_date = None

    with pdfplumber.open(pdf_path) as pdf:

        for page in pdf.pages:

            text = page.extract_text()
            if not text:
                continue

            lines = text.split("\n")

            
            # -------- HEADER PARSING --------

            po_match = re.search(r"PO\s*#\s*:\s*(\d+)", text)
            if po_match:
                current_po = po_match.group(1)

            # Extract ALL dates on page
            all_dates = re.findall(r"\d{2}/\d{2}/\d{4}", text)

            if all_dates:
                # First date after PO block = Order Date
                order_date = all_dates[0]

                # Second date usually = Delivery Date
                if len(all_dates) > 1:
                    delivery_date = all_dates[1]

            # ---------------- ITEM PARSING ----------------

            i = 0
            while i < len(lines):

                line = lines[i].strip()

                # Must start with item number
                if not re.match(r"^\d+\s+", line):
                    i += 1
                    continue

                # Must contain price pattern
                numbers = re.findall(r"[\d,]+\.\d+", line)
                if len(numbers) < 2:
                    i += 1
                    continue

                price_per_unit = float(numbers[0].replace(",", ""))
                total_price = float(numbers[-1].replace(",", ""))

                # Extract quantity + unit (EA / BOX / etc)
                qty_match = re.search(r"/\d+\s+(\w+)\s+(\d+)\s+\w+", line)
                if not qty_match:
                    i += 1
                    continue

                unit = qty_match.group(1)
                qty = float(qty_match.group(2))

                # -------- Extract product code --------

                left_part = line.split(str(numbers[0]))[0].strip()
                parts = left_part.split()

                # remove item number
                parts = parts[1:]

                # remove [M] or [S]
                parts = [p for p in parts if not re.match(r"\[.*\]", p)]

                product_code = None
                for p in reversed(parts):
                    if p.isdigit():
                        product_code = p
                        break

                if not product_code:
                    i += 1
                    continue

                # -------- Extract Thai product name --------

                product_name = None

                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()

                    dc_match = re.match(r"^DC\d+\s+(.+)", next_line)
                    if dc_match:
                        thai_full = dc_match.group(1)

                        # remove (1:1) FT etc
                        thai_clean = re.split(r"\(", thai_full)[0].strip()
                        product_name = thai_clean

                if not product_name:
                    product_name = "Unknown"

                # -------- Save row --------

                items.append({
                    "PO Number": current_po,
                    "Order Date": order_date,
                    "Delivery Date": delivery_date,
                    "Product Name": product_name,
                    "Product Code": product_code,
                    "Unit": unit,
                    "Price per Unit": price_per_unit,
                    "Qty": qty,
                    "Total Price (Ex VAT)": total_price
                })

                i += 2  # skip DC line
            # end while

    df = pd.DataFrame(items)

    print("Rows extracted:", len(df))

    return df

def file_over_file(pdf_path):

    path_dir = os.path.dirname(pdf_path)

    all_dfs = []

    for file in os.listdir(path_dir):

        if file.lower().endswith(".pdf"):   # avoid non-pdf files
            full_path = os.path.join(path_dir, file)

            print("Processing:", file)

            df = get_data_homepro(full_path)

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
    # pdf_path = r"Z:\01_DATA\po\HomePro\PO_0000003074_20251015_083114533353.PDF"
    pdf_path = r"Z:\01_DATA\po\HomePro-New\PO_0000003074_20260225_091029173970.PDF"
    df = get_data_homepro(pdf_path)





    # df = file_over_file(pdf_path)

    if not df.empty:
        path_dir = os.path.dirname(pdf_path)
        file_name = os.path.basename(pdf_path)
        file_name_no_ext = os.path.splitext(file_name)[0]

        output_path = os.path.join(path_dir, f"Extracted_{file_name_no_ext}.xlsx")
        df.to_excel(output_path, index=False)
        print("Saved to:", output_path)
    else:
        print("No data extracted.")
