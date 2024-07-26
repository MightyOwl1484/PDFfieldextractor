import fitz  # PyMuPDF
import openpyxl
import re

def extract_pdf_fields(pdf_path):
    # Open the PDF file
    document = fitz.open(pdf_path)
    
    # Check if the document is encrypted
    if document.is_encrypted:
        # Attempt to decrypt the document
        if document.authenticate(""):
            print("Document decrypted successfully.")
        else:
            print("Failed to decrypt the document. Text extraction may be restricted.")
            return []
    
    # Extract field names and values
    data = []
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        for field in page.widgets():
            if field.field_name:
                data.append((field.field_name, field.field_value))
    
    print(f"Extracted {len(data)} fields from the PDF.")
    return data

def filter_and_extract_six_digit_integers(fields):
    buno_pattern = re.compile(r'11 BUNO Aircraft BUNO Sighted\d+')
    six_digit_integers = []

    for name, value in fields:
        if buno_pattern.match(name) and value and value.isdigit() and len(value) == 6:
            print(f"Field name: {name}, Field value: {value}")
            six_digit_integers.append(value)
    
    print(f"Extracted {len(six_digit_integers)} six-digit integers.")
    return six_digit_integers

def write_data_to_excel(data, excel_path):
    # Create a new Excel workbook and sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Extracted Data"
    
    # Write the header
    sheet.append(["Six Digit Integer"])
    
    # Write the data to the sheet
    for item in data:
        sheet.append([item])
    
    # Save the workbook
    workbook.save(excel_path)
    print(f"Data has been successfully written to {excel_path}")

def main():
    pdf_path = 'C:/Users/dave2/Downloads/WE5 FY24 Aircraft W2W Count Sheet (AMARG) June2024.pdf'  # Update this with the path to your PDF file
    excel_path = 'C:/Users/dave2/OneDrive/Desktop/amage output.xlsx'  # Update this with the desired path to your Excel file
    
    fields = extract_pdf_fields(pdf_path)
    
    # Print a few sample fields for inspection
    for i, (name, value) in enumerate(fields):
        if i < 20:  # Print first 20 fields for inspection
            print(f"Sample field {i+1}: Field name: {name}, Field value: {value}")
    
    if fields:
        six_digit_integers = filter_and_extract_six_digit_integers(fields)
        if six_digit_integers:
            write_data_to_excel(six_digit_integers, excel_path)
        else:
            print("No six-digit integers found.")
    else:
        print("No data extracted. Please check the PDF file and permissions.")

if __name__ == "__main__":
    main()
