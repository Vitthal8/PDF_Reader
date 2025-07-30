import re
from PyPDF2 import PdfReader
import os
import pandas as pd
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)

def extract_records_from_pdf(pdf_path):
    records = []

    with open(pdf_path, "rb") as file:
        pdf_reader = PdfReader(file)

        ref_number = None
        date = None
        customer_name = None
        customer_address = []
        is_collecting_address = False

        for page_index, page in enumerate(pdf_reader.pages):
            text = page.extract_text()
            logging.info(f'Extracted text from page {page_index + 1}: {text}')

            if text is None:
                logging.warning(f'No text found on page {page_index + 1}')
                continue

            lines = text.splitlines()

            for line in lines:
                line = line.strip()
                logging.info(f'Processing line: {line}')

                # Capture Reference Number and Date
                if re.match(r"^Ref\.:.*Date:", line, re.IGNORECASE):
                    ref_match = re.search(r"Ref\.\s*:\s*(.+?)\s*Date\s*:\s*(.+)", line, re.IGNORECASE)
                    if ref_match:
                        ref_number = ref_match.group(1).strip()
                        date = ref_match.group(2).strip()
                        is_collecting_address = True
                        logging.info(f'Found Reference No: {ref_number}, Date: {date}')

                elif is_collecting_address:
                    if customer_name is None:
                        customer_name = line
                        logging.info(f'Customer Name: {customer_name}')

                    if "Subject :" in line:
                        is_collecting_address = False
                        if all([ref_number, date, customer_name]):
                            records.append({
                                "Reference No": ref_number,
                                "Date": date,
                                "Customer Name": customer_name.strip(),
                                "Customer Address": ' '.join(customer_address).strip(),
                                "PDF Filename": os.path.basename(pdf_path),
                            })
                            logging.info('Added record to list')
                        ref_number, date, customer_name = None, None, None
                        customer_address = []
                        continue

                    if line != customer_name:
                        customer_address.append(line)

        # Add last record if exists
        if all([ref_number, date, customer_name]):
            records.append({
                "Reference No": ref_number,
                "Date": date,
                "Customer Name": customer_name.strip(),
                "Customer Address": ' '.join(customer_address).strip(),
                "PDF Filename": os.path.basename(pdf_path),
            })
            logging.info('Added last record to list')

    return records

def create_mis_file(records, output_excel_path):
    df = pd.DataFrame(records)

    # Extract PIN from 'Customer Address' for each record
    for record in records:
        address = record.get('Customer Address', '')
        pin_match = re.search(r'\b(\d{6})\b', address)
        record['PIN'] = pin_match.group(1) if pin_match else ''

    if not df.empty:
        # Insert 'SL' column as first (serial number)
        #df.insert(0, 'SL', range(1, len(df) + 1))

        # Reorder columns so 'PIN' is after 'Customer Address'
        cols = list(df.columns)
        if 'Customer Address' in cols and 'PIN' in cols:
            idx = cols.index('Customer Address')
            # Move 'PIN' to right after 'Customer Address'
            cols.insert(idx + 1, cols.pop(cols.index('PIN')))
            df = df[cols]

        # Save to Excel
        df.to_excel(output_excel_path, index=False)
        logging.info(f'MIS Excel file created at: {output_excel_path}')
    else:
        logging.warning('No records to write to Excel.')

if __name__ == "__main__":
    input_directory = "C:/Project/PDF Reader STK (ACOP)/Input"  # Your folder for PDFs
    output_excel_path = "C:/Project/PDF Reader STK (ACOP)/Output/MIS_File.xlsx"  # Output path

    all_records = []

    logging.info("Starting PDF processing")
    for filename in os.listdir(input_directory):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_directory, filename)
            logging.info(f"Processing file: {pdf_path}")

            # Reset SL counter for each PDF
            sl_counter = 1

            # Extract records from current PDF
            records = extract_records_from_pdf(pdf_path)

            # Assign serial number (SL)
            for record in records:
                record['SL'] = sl_counter
                sl_counter += 1

            all_records.extend(records)

    # Post-process all records: extract PINs
    for record in all_records:
        address = record.get('Customer Address', '')
        pin_match = re.search(r'\b(\d{6})\b', address)
        record['PIN'] = pin_match.group(1) if pin_match else ''

    # Generate Excel file
    create_mis_file(all_records, output_excel_path)
    logging.info("PDF processing completed.")