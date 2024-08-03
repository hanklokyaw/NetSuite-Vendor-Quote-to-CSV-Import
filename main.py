import camelot
import pandas as pd
import os

def is_valid_pdf(file_path):
    return file_path.lower().endswith('.pdf')

# Get the file path from the user
filepath = input('Enter your file path: ')

# Remove surrounding quotes if present
if filepath.startswith('"') and filepath.endswith('"'):
    filepath = filepath[1:-1]

# Replace backslashes with forward slashes
filepath = filepath.replace('\\', '/')
print(filepath)

# Check if the file path is valid
if not is_valid_pdf(filepath):
    print("File path is not a valid PDF.")
else:
    def convert_pdf_to_excel(source_file):
        # Get the base name and directory of the source file
        base_name = os.path.basename(source_file)
        directory = os.path.dirname(source_file)

        # Replace .pdf with .xlsx
        excel_file = os.path.join(directory, base_name.replace('.pdf', '.xlsx'))

        # Read PDF file using camelot
        tables = camelot.read_pdf(source_file, pages='all', flavor='stream')

        # Save each table to an Excel file
        with pd.ExcelWriter(excel_file) as writer:
            for i, table in enumerate(tables):
                table.df.to_excel(writer, sheet_name=f"Sheet_{i}", index=False)

        print(f"Converted {source_file} to {excel_file}")

    # Example usage
    convert_pdf_to_excel(filepath)
    print("Successfully Converted!")
