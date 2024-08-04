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

def check_orderd_string(df):
    for col in df.columns:
        # Check if "Ordered" is in the column
        if df[col].str.contains('Ordered', case=False, na=False).any():
            # Get the index of the "Ordered" row
            ordered_index = df[df[col].str.contains('Ordered', case=False, na=False)].index[0]

            # Iterate through the rows after the "Ordered" row
            for i in range(ordered_index + 1, df.shape[0]):
                try:
                    # Attempt to convert the value to float
                    value = float(df[col].iloc[i])
                    # If successful and the value is not NaN, return True
                    if not pd.isnull(value):
                        return True
                except (ValueError, TypeError):
                    # Ignore errors from non-numeric values
                    continue
            # If no numeric value is found, return False
            return False
        # If "Ordered" is not found in any column, return False
    return False


def transform_data(df):
    """
    Transform the input DataFrame into the desired format.

    Parameters:
    df (pd.DataFrame): The DataFrame containing the raw data.

    Returns:
    pd.DataFrame: The transformed DataFrame.
    """
    # Find the column index for "Ordered"
    ordered_col = df.apply(lambda col: col.str.contains('Ordered', case=False, na=False).any()).idxmax()

    # Find the column index for "Item ID"
    item_id_col = df.apply(lambda col: col.str.contains('Item ID', case=False, na=False).any()).idxmax()

    # Find the column index for the first "Unit" followed by "Price"
    unit_col = df.apply(lambda col: col.str.contains('Unit', case=False, na=False).any()).idxmax()
    price_col = df.apply(lambda col: col.str.contains('Price', case=False, na=False).any()).idxmax()

    # Ensure the "Price" column is to the right of the "Unit" column
    price_col = max(unit_col, price_col)

    # Extract relevant data starting from the first non-NaN row after "Ordered"
    ordered_start_index = df[df.iloc[:, ordered_col].notna()].index[0]
    values_start_index = ordered_start_index + 1

    # Extract the relevant columns and rows
    transformed_df = df.iloc[values_start_index:, [ordered_col, item_id_col, price_col]]

    # Remove rows where the "Ordered" column is NaN or ""
    transformed_df = transformed_df[transformed_df.iloc[:, 0].notna()]
    transformed_df = transformed_df[transformed_df.iloc[:, 0] != ""]

    # Rename columns to match the desired output
    transformed_df.columns = ['Ordered', 'Item ID', 'Price']

    # Reset the index and drop the old index column
    transformed_df.reset_index(drop=True, inplace=True)

    # Check if there are NaN values in the 'Item ID' column
    if transformed_df['Item ID'].isna().any():
        transformed_df = transformed_df.dropna(subset=['Item ID'])

    return transformed_df


# def netsuite_import_sku(filepath):
#     excel_path = filepath.replace(".pdf", ".xlsx")
#     df = pd.read_excel(excel_path, sheet_name="Sheet_2")
#     if check_orderd_string(df):
#         transformed_df = transform_data(df)
#         print(transformed_df)
#     else:
#         print("No item in this page.")

def netsuite_import_sku(filepath):

    # Replace the file extension from .pdf to .xlsx
    excel_path = filepath.replace(".pdf", ".xlsx")

    # Read all sheets into a dictionary of DataFrames
    all_sheets = pd.read_excel(excel_path, sheet_name=None)

    # List to hold transformed DataFrames
    transformed_dfs = []

    # Iterate through each sheet
    for sheet_name, df in all_sheets.items():
        # Check if the sheet contains the necessary conditions
        if check_orderd_string(df):
            transformed_df = transform_data(df)
            transformed_dfs.append(transformed_df)
        else:
            print(f"No valid items in sheet '{sheet_name}'.")

    # Concatenate all transformed DataFrames
    if transformed_dfs:
        final_df = pd.concat(transformed_dfs, ignore_index=True)

        # Calculate the subtotal (sum of Ordered * Price)
        # Ensure columns "Ordered" and "Price" are numeric
        final_df["Ordered"] = pd.to_numeric(final_df["Ordered"], errors='coerce')
        final_df["Price"] = pd.to_numeric(final_df["Price"], errors='coerce')
        final_df["Total"] = final_df["Ordered"] * final_df["Price"]

        # Calculate subtotal
        subtotal = final_df["Total"].sum()
        print(f"Sub-total Amount: {subtotal:.2f}")

        final_df = final_df.drop(columns=['Total'])
    else:
        final_df = pd.DataFrame()
        print("No valid items found in any sheets.")

    print(final_df)

netsuite_import_sku(filepath)