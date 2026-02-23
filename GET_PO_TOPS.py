import pdfplumber
import pandas as pd
import re
import sys
import os


# def get_tops_data(pdf_path, excel_path):

#     data = []
#     current_po = None
#     current_date = None
#     first_item_in_po = True

#     po_pattern = re.compile(r"เลขที่ใบสั่งซื้อ\s*:\s*\n?\s*(\d{6,})")
#     date_pattern = re.compile(r"วันที่.*?:\s*\n?\s*(\d{2}/\d{2}/\d{4})")

#     # Item line:
#     # 1 4897024510198 description ... unit ... qty price discount total
#     item_pattern = re.compile(
#         r"^\s*\d+\s+(\d{8,13})\s+(.*?)\s+(ชิ้น|แพค|ขวด|กล่อง|แพ็ค).*?"
#         r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"
#         r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"
#         r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"
#         r"(\d+(?:,\d+)?(?:\.\d+)?)",
#         re.MULTILINE
#     )

#     with pdfplumber.open(pdf_path) as pdf:
#         for page in pdf.pages:

#             text = page.extract_text()
#             if not text:
#                 continue

#             # Detect new PO
#             po_match = po_pattern.search(text)
#             if po_match:
#                 current_po = po_match.group(1)
#                 first_item_in_po = True

#             # Detect date
#             date_match = date_pattern.search(text)
#             if date_match:
#                 current_date = date_match.group(1)

#             # Extract items
#             for match in item_pattern.finditer(text):

#                 item_code = match.group(1)
#                 description = match.group(2).strip()
#                 unit = match.group(3)
#                 qty = match.group(4).replace(",", "")
#                 unit_price = match.group(5).replace(",", "")
#                 price_after_discount = match.group(6).replace(",", "")
#                 total_price = match.group(7).replace(",", "")

#                 if first_item_in_po:
#                     po_value = current_po
#                     date_value = current_date
#                     first_item_in_po = False
#                 else:
#                     po_value = ""
#                     date_value = ""

#                 data.append([
#                     po_value,
#                     date_value,
#                     item_code,
#                     description,
#                     unit,
#                     qty,
#                     unit_price,
#                     price_after_discount,
#                     total_price
#                 ])

#     df = pd.DataFrame(data, columns=[
#         "Invoice Number",
#         "Order Date",
#         "Item Code",
#         "Description",
#         "Unit",
#         "Qty",
#         "Unit Price",
#         "Price After Discount",
#         "Total Price"
#     ])

#     df.to_excel(excel_path, index=False)

#     print("Extraction completed.")

# ================================
# def get_tops_data(pdf_path, excel_path):

#     data = []
#     current_po = None
#     current_date = None
#     first_item_in_po = True

#     # ---- PO + Date patterns ----
#     po_pattern = re.compile(
#         r"เลขที่ใบสั่งซื้อ\s*:\s*\n?\s*(\d{6,})"
#     )

#     date_pattern = re.compile(
#         r"วันที่.*?:\s*\n?\s*(\d{2}/\d{2}/\d{4})"
#     )

#     # ---- Item pattern ----
#     # Structure:
#     # running_no item_code description unit
#     # pieces_per_pack unit_per_pack qty_pieces price_after_discount total_price

#     item_pattern = re.compile(
#         r"^\s*\d+\s+(\d{8,13})\s+"          # item code
#         r"(.*?)\s+"                        # description
#         r"(ชิ้น|แพค|แพ็ค|ขวด|กล่อง).*?"   # unit
#         r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"     # pieces per pack
#         r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"     # unit per pack
#         r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"     # qty pieces
#         r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"     # price after discount
#         r"(\d+(?:,\d+)?(?:\.\d+)?)",       # total price
#         re.MULTILINE
#     )

#     with pdfplumber.open(pdf_path) as pdf:
#         for page in pdf.pages:

#             text = page.extract_text()
#             if not text:
#                 continue

#             # ---- Detect PO ----
#             po_match = po_pattern.search(text)
#             if po_match:
#                 current_po = po_match.group(1)
#                 first_item_in_po = True

#             # ---- Detect Date ----
#             date_match = date_pattern.search(text)
#             if date_match:
#                 current_date = date_match.group(1)

#             # ---- Extract items ----
#             for match in item_pattern.finditer(text):

#                 item_code = match.group(1)
#                 description = match.group(2).strip()
#                 unit = match.group(3)

#                 pieces_per_pack = match.group(4).replace(",", "")
#                 unit_per_pack = match.group(5).replace(",", "")
#                 qty_pieces = match.group(6).replace(",", "")
#                 price_after_discount = match.group(7).replace(",", "")
#                 total_price = match.group(8).replace(",", "")

#                 if first_item_in_po:
#                     po_value = current_po
#                     date_value = current_date
#                     first_item_in_po = False
#                 else:
#                     po_value = ""
#                     date_value = ""

#                 data.append([
#                     po_value,
#                     date_value,
#                     item_code,
#                     description,
#                     unit,
#                     pieces_per_pack,
#                     unit_per_pack,
#                     qty_pieces,
#                     price_after_discount,
#                     total_price
#                 ])

#     df = pd.DataFrame(data, columns=[
#         "Invoice Number",
#         "Order Date",
#         "Item Code",
#         "Description",
#         "Unit",
#         "Pieces per Pack",
#         "Unit per Pack",
#         "Qty Pieces",
#         "Price After Discount",
#         "Total Price"
#     ])

#     df.to_excel(excel_path, index=False)

#     print("Extraction completed.")

def get_tops_data(pdf_path):

    data = []
    current_po = None
    current_date = None
    first_item_in_po = True

    # Item pattern (same as before)
    item_pattern = re.compile(
        r"^\s*\d+\s+(\d{8,13})\s+"
        r"(.*?)\s+"
        r"(ชิ้น|แพค|แพ็ค|ขวด|กล่อง).*?"
        r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"
        r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"
        r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"
        r"(\d+(?:,\d+)?(?:\.\d+)?)\s+"
        r"(\d+(?:,\d+)?(?:\.\d+)?)",
        re.MULTILINE
    )

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            # --------------------------
            # 1️⃣ Extract PO + Date from top-right area
            # --------------------------

            width = page.width
            height = page.height

            # Crop right upper 30% width, 25% height
            header_crop = page.crop((width * 0.6, 0, width, height * 0.25))
            header_text = header_crop.extract_text()

            if header_text:
                po_match = re.search(r"\b\d{6,}\b", header_text)
                if po_match:
                    current_po = po_match.group(0)
                    first_item_in_po = True

                date_match = re.search(r"\d{2}/\d{2}/\d{4}", header_text)
                if date_match:
                    current_date = date_match.group(0)

            # --------------------------
            # 2️⃣ Extract full page text for items
            # --------------------------

            text = page.extract_text()
            if not text:
                continue

            # --------------------------
            # 3️⃣ Extract items
            # --------------------------

            for match in item_pattern.finditer(text):

                item_code = match.group(1)
                description = match.group(2).strip()
                unit = match.group(3)

                pieces_per_pack = match.group(4).replace(",", "")
                unit_per_pack = match.group(5).replace(",", "")
                qty_pieces = match.group(6).replace(",", "")
                price_after_discount = match.group(7).replace(",", "")
                total_price = match.group(8).replace(",", "")

                if first_item_in_po:
                    po_value = current_po
                    date_value = current_date
                    first_item_in_po = False
                else:
                    po_value = ""
                    date_value = ""

                data.append([
                    po_value,
                    date_value,
                    item_code,
                    description,
                    unit,
                    pieces_per_pack,
                    unit_per_pack,
                    qty_pieces,
                    price_after_discount,
                    total_price
                ])

    df = pd.DataFrame(data, columns=[
        "Invoice Number",
        "Order Date",
        "Item Code",
        "Description",
        "Unit",
        "Pieces per Pack",
        "Unit per Pack",
        "Qty Pieces",
        "Price After Discount",
        "Total Price"
    ])

    path_dir = os.path.dirname(pdf_path)
    file_name = os.path.basename(pdf_path)
    file_name_no_ext = os.path.splitext(file_name)[0]

    df.to_excel(os.path.join(path_dir, f"Extracted_{file_name_no_ext}.xlsx"), index=False)

    print("Extraction completed.")

if __name__ == "__main__":
    # File paths
    
    pdf_path = r"C:\Users\topde\Downloads\tops\Tops\Tops1.pdf"
    excel_path = r"C:\Users\topde\Downloads\tops\Extracted_Tops_Orders-1.xlsx"

    path_dir = os.path.dirname(pdf_path)

    for i in os.listdir(path_dir):
        print(i)
        get_tops_data(os.path.join(path_dir, i))
   
