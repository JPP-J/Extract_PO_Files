import pdfplumber
import pandas as pd
import re
import os


def get_data_lotus(pdf_path):

    header_data = {
        "PO Number": None,
        "Order Date": None,
        "Delivery Date": None,
        "Cancelled Date": None
    }

    items = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            text = page.extract_text()
            if not text:
                continue

            # Clean CID garbage
            text = re.sub(r"\(cid:\d+\)", "", text)

        
            # -----------------------------------
            # Header extraction (encoding-safe)
            # -----------------------------------

            # PO Number (this usually still works)
            po_match = re.search(r"เลข.*?:\s*(\d+)", text)
            if po_match:
                header_data["PO Number"] = po_match.group(1)

            # Extract ALL dates (with optional time)
            all_dates = re.findall(
                r"\d{2}/\d{2}/\d{4}(?:\s+\d{2}:\d{2})?",
                text
            )

            if len(all_dates) >= 1:
                header_data["Order Date"] = all_dates[0]

            if len(all_dates) >= 2:
                header_data["Delivery Date"] = all_dates[1]

            if len(all_dates) >= 3:
                header_data["Cancelled Date"] = all_dates[2]

            # -----------------------------
            # Item extraction
            # -----------------------------
            lines = text.split("\n")
            i = 0

            while i < len(lines):

                line = lines[i].strip()

                # Match first product line
                # Starts with running number then product code
                if re.match(r"^\d+\s+\d+_", line):

                    # Extract barcode
                    barcode_match = re.search(r"\((\d+)/", line)
                    barcode = barcode_match.group(1) if barcode_match else None

                    # Extract all numeric values at end
                    numbers = re.findall(r"\d+\.\d+", line)

                    if len(numbers) >= 6:
                        qty = float(numbers[0])
                        free_gift = float(numbers[1])
                        total_qty = float(numbers[2])
                        price_per_unit = float(numbers[3])
                        discount = float(numbers[4])
                        total_price = float(numbers[5])
                    else:
                        i += 1
                        continue

                    # Extract unit and VAT
                    unit_match = re.search(r"\d+\s+(\S+)\s+([A-Z])\s+\(", line)
                    unit = None
                    vat = None
                    if unit_match:
                        unit = unit_match.group(1)
                        vat = unit_match.group(2)

                    # Extract numeric unit count (before unit text)
                    unit_count_match = re.search(r"_\s+(\d+)\s+", line)
                    unit_count = unit_count_match.group(1) if unit_count_match else None

                    # ----------------------
                    # Product Name (next lines)
                    # ----------------------
                    name_lines = []

                    j = i + 1
                    while j < len(lines):

                        next_line = lines[j].strip()

                        if next_line.startswith("DEPT:"):
                            break

                        if re.match(r"^\d+\s+\d+_", next_line):
                            break

                        name_lines.append(next_line)
                        j += 1

                    product_name = " ".join(name_lines).strip()

                    items.append({
                        "PO Number": header_data["PO Number"],
                        "Order Date": header_data["Order Date"],
                        "Delivery Date": header_data["Delivery Date"],
                        "Cancelled Date": header_data["Cancelled Date"],
                        "Product Name": product_name,
                        "Unit Count": unit_count,
                        "Unit": unit,
                        "VAT": vat,
                        "Product Barcode": barcode,
                        "Qty": qty,
                        "Free Gift": free_gift,
                        "Total Qty": total_qty,
                        "Price per Unit": price_per_unit,
                        "Discount": discount,
                        "Total Price": total_price
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
    pdf_path = r"C:\Users\topde\Downloads\tops\lotus\lotus1.pdf"
    # get_data_lotus(pdf_path)

    path_dir = os.path.dirname(pdf_path)

    for i in os.listdir(path_dir):
        print(i)
        get_data_lotus(os.path.join(path_dir, i))
   