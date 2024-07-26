# PDFfieldextractor
Aircraft Inventory PDF Field Extractor
This Python script extracts specific six-digit integers from fields in a PDF file and writes them to an Excel file. It uses the PyMuPDF and openpyxl libraries to achieve this.

Prerequisites
Before running the script, ensure you have the following Python libraries installed:

fitz (PyMuPDF)
openpyxl
re (standard library)
You can install the required libraries using pip:

bash
Copy code
pip install pymupdf openpyxl
Script Overview
The script consists of the following functions:

extract_pdf_fields(pdf_path): Opens the PDF file, extracts field names and values, and handles encrypted documents.
filter_and_extract_six_digit_integers(fields): Filters the extracted fields to find specific six-digit integers.
write_data_to_excel(data, excel_path): Writes the filtered six-digit integers to an Excel file.
main(): Main function that orchestrates the extraction, filtering, and writing processes.
Usage
Update the pdf_path and excel_path variables in the main() function with the paths to your PDF file and desired Excel output file, respectively.
Run the script.
Detailed Function Description
extract_pdf_fields(pdf_path)
Parameters:
pdf_path (str): Path to the PDF file.
Returns:
data (list): A list of tuples containing field names and values.
This function opens the specified PDF file, checks if it is encrypted, and attempts to decrypt it if necessary. It then extracts all field names and their values from the PDF.

filter_and_extract_six_digit_integers(fields)
Parameters:
fields (list): A list of tuples containing field names and values.
Returns:
six_digit_integers (list): A list of six-digit integers extracted from the fields.
This function filters the extracted fields to find specific six-digit integers based on a regex pattern matching certain field names.

write_data_to_excel(data, excel_path)
Parameters:
data (list): A list of six-digit integers to be written to the Excel file.
excel_path (str): Path to the desired Excel file.
Returns: None
This function creates a new Excel workbook, writes the extracted six-digit integers to it, and saves the workbook at the specified path.

main()
The main function that orchestrates the workflow:

Extracts fields from the PDF.
Filters and extracts six-digit integers.
Writes the extracted data to an Excel file.
Example
Extracting Fields from PDF: The script opens the specified PDF file and extracts fields.
Filtering Six-Digit Integers: Filters the fields to find six-digit integers that match a specific pattern.
Writing to Excel: Writes the filtered integers to an Excel file.
python
Copy code
if __name__ == "__main__":
    main()
Notes
Ensure the PDF file is not encrypted or can be decrypted by the script for successful extraction.
Modify the pdf_path and excel_path variables in the main() function according to your file paths.
