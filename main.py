import camelot
import pandas as pd
import os
from datetime import datetime

# Define filenames with today's date
today_date = datetime.today().strftime('%Y%m%d')
formatted_today_date = datetime.today().strftime('%m/%d/%Y')
sku_filename = f'item_sku_{today_date}_tooltech.csv'
po_filename = f'ns_non-inv_so_{today_date}_tooltech.csv'

vendor_name = '4062 Tool Technology Distributors, Inc.'
purchaser = "Hank Kyaw"

def is_valid_pdf(file_path):
    """
    Check if the provided file path points to a PDF file.
    """
    return file_path.lower().endswith('.pdf')

def convert_pdf_to_excel(source_file):
    """
    Convert a PDF file to an Excel file by extracting tables.
    """
    base_name = os.path.basename(source_file)
    directory = os.path.dirname(source_file)
    excel_file = os.path.join(directory, base_name.replace('.pdf', '.xlsx'))

    # Read PDF file using camelot
    tables = camelot.read_pdf(source_file, pages='all', flavor='stream')

    # Save each table to an Excel file
    with pd.ExcelWriter(excel_file) as writer:
        for i, table in enumerate(tables):
            table.df.to_excel(writer, sheet_name=f"Sheet_{i}", index=False)

    print(f"Converted {source_file} to {excel_file}")

def filepath_to_excel(iteration):
    """
    Get the file path, PO ID, and memo from the user, validate the file, and convert it if valid.
    """
    filepath = input(f'Enter your filepath {iteration}: ')
    po_id = input(f'Enter your Ana PO id {iteration}: ')
    memo = input(f'Enter your Memo {iteration}: ')

    if filepath.startswith('"') and filepath.endswith('"'):
        filepath = filepath[1:-1]

    filepath = filepath.replace('\\', '/')
    print(filepath)

    if not is_valid_pdf(filepath):
        print("File path is not a valid PDF.")
    else:
        convert_pdf_to_excel(filepath)
        print("Successfully Converted!")

    return filepath, po_id, memo

def check_orderd_string(df):
    """
    Check if any column in the DataFrame contains "Ordered" followed by numeric values.
    """
    for col in df.columns:
        if df[col].str.contains('Ordered', case=False, na=False).any():
            ordered_index = df[df[col].str.contains('Ordered', case=False, na=False)].index[0]
            for i in range(ordered_index + 1, df.shape[0]):
                try:
                    value = float(df[col].iloc[i])
                    if not pd.isnull(value):
                        return True
                except (ValueError, TypeError):
                    continue
    return False

def transform_data(df):
    """
    Transform the input DataFrame into a desired format.
    """
    ordered_col = df.apply(lambda col: col.str.contains('Ordered', case=False, na=False).any()).idxmax()
    item_id_col = df.apply(lambda col: col.str.contains('Item ID', case=False, na=False).any()).idxmax()
    unit_col = df.apply(lambda col: col.str.contains('Unit', case=False, na=False).any()).idxmax()
    price_col = df.apply(lambda col: col.str.contains('Price', case=False, na=False).any()).idxmax()

    price_col = max(unit_col, price_col)
    ordered_start_index = df[df.iloc[:, ordered_col].notna()].index[0]
    values_start_index = ordered_start_index + 1

    transformed_df = df.iloc[values_start_index:, [ordered_col, item_id_col, price_col]]
    transformed_df = transformed_df[transformed_df.iloc[:, 0].notna()]
    transformed_df = transformed_df[transformed_df.iloc[:, 0] != ""]
    transformed_df.columns = ['Ordered', 'Item ID', 'Price']
    transformed_df.reset_index(drop=True, inplace=True)

    if transformed_df['Item ID'].isna().any():
        transformed_df = transformed_df.dropna(subset=['Item ID'])

    return transformed_df

def netsuite_import_sku(filepath):
    """
    Import SKU data from the Excel file created from a PDF and transform it.
    """
    excel_path = filepath.replace(".pdf", ".xlsx")
    all_sheets = pd.read_excel(excel_path, sheet_name=None)
    transformed_dfs = []

    for sheet_name, df in all_sheets.items():
        if check_orderd_string(df):
            transformed_df = transform_data(df)
            transformed_dfs.append(transformed_df)
        else:
            print(f"No valid items in sheet '{sheet_name}'.")

    if transformed_dfs:
        final_df = pd.concat(transformed_dfs, ignore_index=True)
        final_df["Ordered"] = pd.to_numeric(final_df["Ordered"], errors='coerce')
        final_df["Price"] = pd.to_numeric(final_df["Price"], errors='coerce')
        final_df["Total"] = final_df["Ordered"] * final_df["Price"]
        subtotal = final_df["Total"].sum()
        print(f"Sub-total Amount: {subtotal:.2f}")

        final_df = final_df.drop(columns=['Total'])
        final_df = final_df.rename(columns={'Item ID':'Item Name'})
        final_df['Vendor'] = vendor_name
        final_df['SKU'] = final_df['Item Name']
        final_df = final_df[['Vendor', 'Item Name', 'SKU', 'Price']]
        return final_df
    else:
        print("No valid items found in any sheets.")
        return pd.DataFrame()

def netsuite_import_so(filepath, po_id, memo):
    """
    Import SO data from the Excel file created from a PDF and transform it.

    Parameters:
    filepath (str): The path to the PDF file.
    po_id (str): The PO ID.
    memo (str): The memo for the SO.

    Returns:
    The transformed SO DataFrame.
    """
    excel_path = filepath.replace(".pdf", ".xlsx")
    all_sheets = pd.read_excel(excel_path, sheet_name=None)
    transformed_dfs = []

    for sheet_name, df in all_sheets.items():
        if check_orderd_string(df):
            transformed_df = transform_data(df)
            transformed_dfs.append(transformed_df)
        else:
            print(f"No valid items in sheet '{sheet_name}'.")

    if transformed_dfs:
        final_df = pd.concat(transformed_dfs, ignore_index=True)
        final_df["Ordered"] = pd.to_numeric(final_df["Ordered"], errors='coerce')
        final_df["Price"] = pd.to_numeric(final_df["Price"], errors='coerce')
        final_df["Total"] = final_df["Ordered"] * final_df["Price"]
        subtotal = final_df["Total"].sum()
        print(f"Sub-total Amount: {subtotal:.2f}")

        final_df = final_df.rename(columns={'Item ID':'Item Name', 'Price':'Rate', 'Total': 'Price', 'Ordered': 'Quantity'})
        final_df['SKU'] = final_df['Item Name']
        final_df['Ana SO'] = po_id
        final_df['Date'] = formatted_today_date
        final_df['Vendor'] = vendor_name
        final_df['Memo'] = memo
        final_df['Purchaser'] = purchaser
        final_df = final_df[['Ana SO', 'Date', 'Vendor', 'Purchaser', 'Memo', 'Item Name', 'SKU', 'Rate', 'Quantity', 'Price']]
        return final_df
    else:
        print("No valid items found in any sheets.")
        return pd.DataFrame()

def extract_integer(input_str):
    """
    Extract an integer from a string, allowing for one decimal point.

    Parameters:
    input_str (str): The input string to extract the integer from.

    Returns:
    int: The extracted integer.
    """
    input_str = input_str.strip()
    num_str = ''
    decimal_found = False

    for char in input_str:
        if char.isdigit():
            num_str += char
        elif char == ',' and not decimal_found:
            decimal_found = True
        elif not char.isdigit():
            break

    try:
        num = int(float(num_str))
    except ValueError:
        num = 0

    return num

def combine_all_sku_and_po():
    """
    Process multiple PDF files to generate SKU and PO CSV files.
    """
    num_of_iteration = input("How many Tool Tech that you need to process? ")
    num_of_iteration = extract_integer(num_of_iteration)

    sku_df_list = []
    po_df_list = []

    for i in range(1, num_of_iteration + 1):
        filepath, po_id, memo = filepath_to_excel(i)

        sku_df = netsuite_import_sku(filepath)
        po_df = netsuite_import_so(filepath, po_id, memo)

        sku_df_list.append(sku_df)
        po_df_list.append(po_df)

    final_sku_df = pd.concat(sku_df_list, ignore_index=True)
    final_po_df = pd.concat(po_df_list, ignore_index=True)

    final_sku_df.to_csv(sku_filename, index=False)
    final_po_df.to_csv(po_filename, index=False)

combine_all_sku_and_po()
