import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os


def convert_date(date_str):
    """
    Convert 'Thu,26 Feb 2026' → '26/02/2026'
    """
    try:
        dt = datetime.strptime(date_str.strip(), "%a,%d %b %Y")
        return dt.strftime("%d/%m/%Y")
    except:
        return None


def get_data_twd(pdf_path):

    rows = []

    header = {
        "PO Number": None,
        "Date": None,
        "Entry Date": None,
        "Ship Date": None
    }

    with pdfplumber.open(pdf_path) as pdf:

        for page in pdf.pages:

            text = page.extract_text()
            if not text:
                continue

            # ---------------- HEADER ----------------

            # Document Date
            date_match = re.search(r"Date\s*:\s*\w+,(\d{1,2}\s+\w+\s+\d{4})", text)
            if date_match and not header["Date"]:
                header["Date"] = convert_date("Thu," + date_match.group(1))

            # PO Number
            po_match = re.search(r"P/O Number\s*:\s*(\d+)", text)
            if po_match and not header["PO Number"]:
                header["PO Number"] = po_match.group(1)

            # Entry Date
            entry_match = re.search(r"Entry Date\s*:\s*\w+,(\d{1,2}\s+\w+\s+\d{4})", text)
            if entry_match and not header["Entry Date"]:
                header["Entry Date"] = convert_date("Wed," + entry_match.group(1))

            # Ship Date
            ship_match = re.search(r"Ship Date\s*:\s*\w+,(\d{1,2}\s+\w+\s+\d{4})", text)
            if ship_match and not header["Ship Date"]:
                header["Ship Date"] = convert_date("Wed," + ship_match.group(1))

            # ---------------- ITEMS ----------------

            lines = text.split("\n")

            for line in lines:

                line = line.strip()

                # Match item line starting with 13-digit barcode
                item_match = re.match(
                    r"^(\d{13})\s+\d+\s+(.+?)\s+\d+\s+\d+\s+\d+\s+EA\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)",
                    line
                )

                if item_match:

                    barcode = item_match.group(1)
                    product_name = item_match.group(2).strip()
                    net_amount = float(item_match.group(3).replace(",", ""))

                    rows.append({
                        "PO Number": header["PO Number"],
                        "Date": header["Date"],
                        "Entry Date": header["Entry Date"],
                        "Ship Date": header["Ship Date"],
                        "Barcode": barcode,
                        "Product Name": product_name,
                        "Net (Ex VAT)": net_amount
                    })

    df = pd.DataFrame(rows)

    print("Rows extracted:", len(df))

    return df


def file_over_file(pdf_path):

    path_dir = os.path.dirname(pdf_path)

    all_dfs = []

    for file in os.listdir(path_dir):

        if file.lower().endswith(".pdf"):   # avoid non-pdf files
            full_path = os.path.join(path_dir, file)

            print("Processing:", file)

            df = get_data_twd(full_path)

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
    pdf_path = r"Z:\01_DATA\po\ไทวัสดุ\PO_Info (1).pdf"
    df = get_data_twd(pdf_path)





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
