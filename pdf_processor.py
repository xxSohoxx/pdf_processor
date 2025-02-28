import os
import pdfplumber
import re
import pandas as pd

UPLOAD_FOLDER = "uploads"
OUTPUT_FILE = "processed_data.xlsx"

def extract_number_from_filename(filename):
    match = re.search(r"(\d+)_Part_", filename)
    return match.group(1) if match else "N/A"

def remove_dimension(value):
    return re.sub(r"\s*[a-zA-Z]+\.?", "", value)

def extract_values_from_pdf(pdf_path):
    keys_to_extract = [
        "Name", "MATERIAL SPECIFICATION", "COATING SPECIFICATION", 
        "MAXIMUM OUTSIDE DIAMETER", "OVERALL LENGTH", "MINIMUM INSIDE DIAMETER"
    ]
    keys_with_dimension = [
        "MAXIMUM OUTSIDE DIAMETER", "OVERALL LENGTH", "MINIMUM INSIDE DIAMETER"
    ]

    extracted_data = {
        "Number": extract_number_from_filename(pdf_path),  # Extract number at the start
        "Name": "N/A"
    }

    extracted_data.update({key: "N/A" for key in keys_to_extract})
    prev_row = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    row = [cell.strip() for cell in row if cell]
                    
                    if not row:  # Skip empty rows
                        continue

                    if row[0] == "Described By Document":
                        continue

                    if prev_row and prev_row[0] == "Basic Attributes":
                        extracted_data["Name"] = row[1]
                    else:
                        for key in keys_to_extract:
                            if key in row:
                                idx = row.index(key)
                                value = row[idx + 1] if idx + 1 < len(row) else "N/A"
                                if key in keys_with_dimension:
                                    extracted_data[key] = remove_dimension(value)
                                else:
                                    extracted_data[key] = value
                    prev_row = row
    print(f'Processed file {pdf_path}.')
    return extracted_data

def process_pdf_folder():
    results = []
    for filename in os.listdir(UPLOAD_FOLDER):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(UPLOAD_FOLDER, filename)
            results.append(extract_values_from_pdf(pdf_path))

    return results

def save_to_excel(data):
    df = pd.DataFrame(data)
    df.rename(columns={
        "Number": "Number",
        "Name": "Name",
        "MATERIAL SPECIFICATION": "WFD Material",
        "COATING SPECIFICATION": "COATING SPECIFICATION",
        "MAXIMUM OUTSIDE DIAMETER": "OD, in",
        "MINIMUM INSIDE DIAMETER": "ID, in",
        "OVERALL LENGTH": "L, in"
    }, inplace=True)

    df.insert(0, "Item", range(1, len(df) + 1))  # Add Item (Position Number)

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Data")
        workbook = writer.book
        worksheet = writer.sheets["Extracted Data"]
        red_format = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        for col_num, col_name in enumerate(df.columns):
            worksheet.set_column(col_num, col_num, 20)
            for row_num, value in enumerate(df[col_name], start=1):
                if value == "N/A":
                    worksheet.write(row_num, col_num, value, red_format)

    print(f"✅ Data saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    extracted_data = process_pdf_folder()
    if extracted_data:
        save_to_excel(extracted_data)
    else:
        print("⚠ No valid PDFs found in uploads.")
