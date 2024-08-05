import camelot
import pandas as pd
import os
from datetime import datetime


today_date = datetime.today().strftime('%Y%m%d')
formatted_today_date = datetime.today().strftime('%m/%d/%Y')
sku_filename =f'item_sku_{today_date}_tooltech.csv'
po_filename = f'ns_non-inv_so_{today_date}_tooltech.csv'

def is_valid_pdf(file_path):
    return file_path.lower().endswith('.pdf')

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

def filepath_to_excel(iteration):
    # Get the file path from the user
    filepath = input(f'Enter your filepath {iteration}: ')
    po_id = input(f'Enter your Ana PO id {iteration}: ')
    memo = input(f'Enter your Memo {iteration}: ')

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
        convert_pdf_to_excel(filepath)
        print("Successfully Converted!")

    return filepath, po_id, memo



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

        # Rearrange SKU import csv columns
        final_df = final_df.drop(columns=['Total'])
        final_df = final_df.rename(columns={'Item ID':'Item Name'})
        final_df['Vendor'] = '4062 Tool Technology Distributors, Inc.'
        final_df['SKU'] = final_df['Item Name']
        final_df = final_df[['Vendor', 'Item Name', 'SKU', 'Price']]
        return final_df
    else:
        final_df = pd.DataFrame()
        print("No valid items found in any sheets.")

    # print(final_df)

def netsuite_import_so(filepath, po_id, memo):

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

        # Rearrange SKU import csv columns
        final_df = final_df.rename(columns={'Item ID':'Item Name',
                                            'Price':'Rate',
                                            'Total': 'Price',
                                            'Ordered': 'Quantity'})
        final_df['SKU'] = final_df['Item Name']
        final_df['Ana SO'] = po_id
        final_df['Date'] = formatted_today_date
        final_df['Vendor'] = '4062 Tool Technology Distributors, Inc.'
        final_df['Memo'] = memo
        final_df['Purchaser'] = "Hank Kyaw"
        final_df = final_df[['Ana SO', 'Date', 'Vendor', 'Purchaser', 'Memo', 'Item Name', 'SKU', 'Rate', 'Quantity', 'Price']]
        return final_df
    else:
        final_df = pd.DataFrame()
        print("No valid items found in any sheets.")

    print(final_df)

def extract_integer(input_str):
    # Trim leading and trailing whitespace
    input_str = input_str.strip()

    # Initialize an empty string to collect numeric characters
    num_str = ''
    decimal_found = False

    for char in input_str:
        # Allow digits and, at most, one decimal point
        if char.isdigit():
            num_str += char
        elif char in (',') and not decimal_found:
            decimal_found = True
        elif not char.isdigit():
            break

    # Convert the extracted numeric string to an integer
    try:
        # Convert to float first, then to integer
        num = int(float(num_str))
    except ValueError:
        # If conversion fails, return 0 (or handle it as per your need)
        num = 0

    return num

def combine_all_sku_and_po():
    num_of_iteration = input("How many Tool Tech that you need to process? ")
    num_of_iteration = extract_integer(num_of_iteration)

    # Initialize empty lists to store DataFrames
    sku_df_list = []
    po_df_list = []

    # Loop over the number of iterations
    for i in range(1, num_of_iteration + 1):
        filepath, po_id, memo = filepath_to_excel(i)

        # Get SKU and PO DataFrames
        sku_df = netsuite_import_sku(filepath)
        po_df = netsuite_import_so(filepath, po_id, memo)

        # Append the DataFrames to the respective lists
        sku_df_list.append(sku_df)
        po_df_list.append(po_df)

    # Concatenate all SKU and PO DataFrames
    final_sku_df = pd.concat(sku_df_list, ignore_index=True)
    final_po_df = pd.concat(po_df_list, ignore_index=True)

    final_sku_df.to_csv(sku_filename, index=False)
    final_po_df.to_csv(po_filename, index=False)

combine_all_sku_and_po()

# netsuite_import_sku(filepath)
# netsuite_import_so(filepath, 24080401, "Weekly points and tips for Swiss Dept.")