from openpyxl import load_workbook
import xlrd
import datetime
import os
from ExcelHelpers import (
    resource_path, FINISHED_FOLDER, format_cells_as_text, align_cells_left, manyToMany, oneToMany, typedValue
)
# Define source files and destination copies for Chewy
source_asn_xlsx = resource_path("assets/Murdochs/Blank Murdochs 856 ASN.xlsx")

# Function to copy data from uploaded .xlsx file to specific cells in the ASN .xlsx backup
def copy_xlsx_data(uploaded_file, dest_file):
    uploaded_wb = load_workbook(uploaded_file)
    source_wb = load_workbook(source_asn_xlsx)
    uploaded_ws = uploaded_wb.active
    source_ws = source_wb.active
    format_cells_as_text(source_ws)
    align_cells_left(source_ws)
    # Mapping uploaded cells to copy cells
    data_map = {
        'B18': 'E3', 'D18': 'E4', 'E18': 'E5',
        'J18': 'E6', 'K18': 'E7', 'L18': 'E8',
        'C10': 'E14'
    }

    # Transfer data from the uploaded file to the backup copy
    for upload_cell, copy_cell in data_map.items():
        source_ws[copy_cell] = uploaded_ws[upload_cell].value

    total_cases = 0  # Initialize a variable to keep track of the total cases

      # Loop through each row to calculate the cases per UPC and accumulate the total cases
    for i in range(23, xls_sheet.nrows):
        upc = xls_sheet.cell_value(i, 5)  # Column F for UPC
        qty = int(xls_sheet.cell_value(i, 8))  # Column I for QTY

        if upc in upc_counts:
            items_per_case = upc_counts[upc]
            cases = qty / items_per_case
            total_cases += cases  # Add to the total cases
            print(f"Row {i + 1}: UPC = {upc}, QTY = {qty}, Items per Case = {items_per_case}, Cases = {cases}")
        else:
            print(f"Warning: UPC {upc} not found in upc_cases dictionary.")
    # Track numbers from A21 and below, copy to A20 in the copy
    row = 21
    data_to_copy = []
    while uploaded_ws[f'A{row}'].value is not None:
        data_to_copy.append(uploaded_ws[f'A{row}'].value)
        row += 1

    for i, value in enumerate(data_to_copy):
        source_ws[f'A{20 + i}'] = value

    count = len(data_to_copy)
    value_from_upload = uploaded_ws['C4'].value

    if value_from_upload is not None:
        for i in range(20, 20 + count):
            source_ws[f'B{i}'] = value_from_upload

    source_wb.save(dest_file)

# Function to convert .xls to .xlsx and transfer data to a backup of ASN copy
def convert_xls_data(uploaded_file, dest_file):
    xls_book = xlrd.open_workbook(uploaded_file)
    source_wb = load_workbook(source_asn_xlsx)
    source_ws = source_wb.active
    xls_sheet = xls_book.sheet_by_index(0)

    # Mapping uploaded cells to copy cells
    data_map = {
        (17, 1): 'E3',   # 'B18' -> (18, 2) name
        (17, 3): 'E4',   # 'D18' -> (18, 4) number
        (17, 4): 'E5',   # 'E18' -> (18, 5) add 1
        (17, 9): 'E6',   # 'J18' -> (18, 10) city
        (17, 10): 'E7',  # 'K18' -> (18, 11) State
        (17, 11): 'E8',  # 'L18' -> (18, 12) Zip
        (9, 2): 'E14'    # 'C10' -> (10, 3) delivery date
    }

    for (row, col), copy_cell in data_map.items():
        value = xls_sheet.cell_value(row, col)
        source_ws[copy_cell] = value

    # Manually assign values for cells that don’t come from the xls_sheet
    source_ws['B12'] = "FEDG"
    source_ws['B13'] = "Fedex"


    # Debugging: Print the total number of rows and columns in the uploaded sheet
    total_rows = xls_sheet.nrows
    total_cols = xls_sheet.ncols
    print(f"Total rows: {total_rows}, Total columns: {total_cols}")

    # Ensure you don't go out of bounds
    if total_rows < 23:
        print(f"Error: Not enough rows to start processing from row 23.")
        return

    # Dynamic column length calculation with boundary check
    row = 23
    column_length = 0

    while row < total_rows:  # Ensure we don't exceed the available rows
        try:
            value = xls_sheet.cell_value(row - 1, 0)  # Column A (index 0)
            print(f"Row {row}: Value in A = {value}")

            if value:
                column_length += 1
                row += 1
            else:
                break  # Stop if an empty cell is found
        except IndexError as e:
            print(f"Error accessing row {row - 1}, column 0: {str(e)}")
            break

    print(f"Dynamic Length of Column A: {column_length}")

    # Use helper function safely with boundary checks
    try:
        if column_length > 0:
            manyToMany(xls_sheet, source_ws, 23, 0, 'A', 19, column_length)  # A23 to A19       Item number
            manyToMany(xls_sheet, source_ws, 23, 5, 'F', 19, column_length)  # F23 to F19       UPC
            manyToMany(xls_sheet, source_ws, 23, 8, 'G', 19, column_length)  # I23 to G19       SKU
            manyToMany(xls_sheet, source_ws, 23, 6, 'H', 19, column_length)  # G23 to H19       Vendor Part
            manyToMany(xls_sheet, source_ws, 23, 1, 'I', 19, column_length)  # B23 to I19       QTY
            manyToMany(xls_sheet, source_ws, 23, 2, 'J', 19, column_length)  # C23 to J19       Unit of Measure
            manyToMany(xls_sheet, source_ws, 23, 4, 'K', 19, column_length)  # E23 to K19 I 23  Description
        else:
            print("No data found in column A starting from row 23.")
    except Exception as e:
        print(f"Error during copy operations: {str(e)}")

    # Additional copy and paste operations with helper functions
    oneToMany(xls_sheet, source_ws, row=3, col=2, target_column='B', start_row=19, column_length=column_length)  # PO
    oneToMany(xls_sheet, source_ws, row=3, col=7, target_column='C', start_row=19, column_length=column_length)  # PO Date

    typedValue(source_ws, static_value="N/A", target_column='D', start_row=19, column_length=column_length)

    format_cells_as_text(source_ws)
    align_cells_left(source_ws)
    align_cells_left(source_ws)

    # Save the updated file
    try:
        source_wb.save(dest_file)
        print(f"File saved successfully as {dest_file}.")
    except Exception as e:
        print(f"Error saving file: {str(e)}")


def process_MurdochsASN(file_path):
    """Main function to process Murdochs ASN files."""
    current_date = datetime.datetime.now().strftime("%m.%d.%Y")

    # Determine if the file is XLSX or XLS and extract the PO number
    if file_path.endswith('.xlsx'):
        uploaded_wb = load_workbook(file_path)
        po_number = uploaded_wb.active['C4'].value
    elif file_path.endswith('.xls'):
        xls_book = xlrd.open_workbook(file_path)
        po_number = xls_book.sheet_by_index(0).cell_value(3, 2)

    # Define the backup file path in FINISHED_FOLDER with a Murdochs subfolder
    backup_file = os.path.join(FINISHED_FOLDER, f"Murdochs/Murdochs 856 ASN PO {po_number} {current_date}.xlsx")
    
    # Ensure the directory exists
    os.makedirs(os.path.dirname(backup_file), exist_ok=True)

    # Perform the copy or conversion based on file type
    if file_path.endswith('.xlsx'):
        copy_xlsx_data(file_path, backup_file)
    elif file_path.endswith('.xls'):
        convert_xls_data(file_path, backup_file)

    return backup_file, po_number
