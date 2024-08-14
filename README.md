# Tool Tech PDF to CSV Conversion

## Overview
This script processes PDF files containing item data and sales orders, converts them to Excel files, and then transforms the data into CSV files suitable for importing into NetSuite.

## Requirements
- Python 3.x
- `camelot-py[cv]` library for PDF table extraction
- `pandas` library for data manipulation

## Script Functions

1. **`is_valid_pdf(file_path)`**
   - Checks if the given file path is a PDF.

2. **`convert_pdf_to_excel(source_file)`**
   - Converts a PDF file to an Excel file by extracting tables.

3. **`filepath_to_excel(iteration)`**
   - Prompts the user to input file path, PO ID, and memo, then converts the file if valid.

4. **`check_orderd_string(df)`**
   - Checks if any column in the DataFrame contains "Ordered" followed by numeric values.

5. **`transform_data(df)`**
   - Transforms the DataFrame into a desired format.

6. **`netsuite_import_sku(filepath)`**
   - Imports SKU data, transforms it, and calculates the subtotal.

7. **`netsuite_import_so(filepath, po_id, memo)`**
   - Imports SO data, transforms it, and calculates the subtotal.

8. **`extract_integer(input_str)`**
   - Extracts an integer from a string, allowing for one decimal point.

9. **`combine_all_sku_and_po()`**
   - Processes multiple PDFs to generate SKU and PO CSV files.

## Usage

1. **Prepare the PDFs:** Ensure your PDF files contain tables formatted with "Ordered", "Item ID", "Unit", and "Price".

2. **Run the Script:**
   - Execute the script in your Python environment.
   - Follow the prompts to input file paths, PO IDs, and memos.
   - The script will process the PDFs, convert them to Excel files, and generate SKU and PO CSV files.

3. **Check Output:**
   - The script will output CSV files named `item_sku_YYYYMMDD_tooltech.csv` and `ns_non-inv_so_YYYYMMDD_tooltech.csv` where `YYYYMMDD` is the current date.

## Notes
- Ensure all required libraries are installed and properly configured.
- Modify the script as needed for different file formats or data requirements.
