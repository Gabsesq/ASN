from openpyxl import load_workbook
import xlrd
import datetime
import os
from ExcelHelpers import (
    resource_path, FINISHED_FOLDER, format_cells_as_text, align_cells_left
)

# Define source files and destination copies for Chewy
source_asn_xlsx = resource_path("assets/Chewy/Chewy 856 ASN - Copy.xlsx")

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
        'B16': 'E3', 'D16': 'E4', 'E16': 'E5',
        'J16': 'E6', 'K16': 'E7', 'L16': 'E8',
        'C8': 'E14'
    }

    # Transfer data from the uploaded file to the backup copy
    for upload_cell, copy_cell in data_map.items():
        source_ws[copy_cell] = uploaded_ws[upload_cell].value

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
    format_cells_as_text(source_ws)
    align_cells_left(source_ws)
    # Mapping uploaded cells to copy cells
    data_map = {
        (15, 1): 'E3', (15, 3): 'E4', (15, 4): 'E5',
        (15, 9): 'E6', (15, 10): 'E7', (15, 11): 'E8',
        (7, 2): 'E14'
    }

    for (row, col), copy_cell in data_map.items():
        value = xls_sheet.cell_value(row, col)
        source_ws[copy_cell] = value

    # Track numbers from A21 and below, copy to A20 in the copy
    row = 21
    copy_row = 20
    column_length = 0

    while True:
        try:
            value = xls_sheet.cell_value(row - 1, 0)  # Column A is index 0 (zero-based)
            if value:
                source_ws[f'A{copy_row}'] = value
                row += 1
                copy_row += 1
                column_length += 1  # Increment the length count
            else:
                break
        except IndexError:
            break

    print(f"Length of Column A: {column_length}")  # Debugging print for column length

    # Now copy from upload file column H (starting at H21) down to E20 in the copy file
    for i in range(21, 21 + column_length):  # Loop through rows in column H in the upload file
        value_from_H = xls_sheet.cell_value(i - 1, 7)  # H21 is (row 20, col 7) in zero-based indexing
        source_ws[f'E{i - 1}'] = value_from_H  # Paste into E20 and down
        print(f"Pasting {value_from_H} from H{i} to E{i - 1}")  # Debugging print

    # Now copy value from C4 into B column (B20 and down for column_length rows)
    value_from_upload_C4 = xls_sheet.cell_value(3, 2)
    if value_from_upload_C4 is not None:
        for i in range(20, 20 + column_length):  # Copy value to B20 through B (dynamic based on column A length)
            source_ws[f'B{i}'] = value_from_upload_C4
            print(f"Pasting {value_from_upload_C4} into B{i}")  # Debugging print
    else:
        print("No value found in C4 or unable to read the value.")  # Debugging if C4 is None

    # Copy date from H4 into C column (C20 and down for column_length rows)
    value_from_upload_H4 = xls_sheet.cell_value(3, 7)
    if value_from_upload_H4 is not None:
        for i in range(20, 20 + column_length):  # Copy value to C20 through C (dynamic based on column A length)
            source_ws[f'C{i}'] = value_from_upload_H4
            print(f"Pasting date {value_from_upload_H4} into C{i}")  # Debugging print
    else:
        print("No date found in H4 or unable to read the value.")  # Debugging if H4 is None

    # Now copy from upload file column G (starting at G21) down to F20 in the copy file
    for i in range(21, 21 + column_length):  # Loop through rows in column G in the upload file
        value_from_G = xls_sheet.cell_value(i - 1, 6)  # G21 is (row 20, col 6) in zero-based indexing
        source_ws[f'F{i - 1}'] = value_from_G  # Paste into F20 and down
        print(f"Pasting {value_from_G} from G{i} to F{i - 1}")  # Debugging print

    # Now copy from upload file column I (starting at I21) down to G20 in the copy file
    for i in range(21, 21 + column_length):  # Loop through rows in column I in the upload file
        value_from_I = xls_sheet.cell_value(i - 1, 8)  # I21 is (row 20, col 8) in zero-based indexing
        source_ws[f'G{i - 1}'] = value_from_I  # Paste into G20 and down
        print(f"Pasting {value_from_I} from I{i} to G{i - 1}")  # Debugging print

    # Now copy from upload file column B (starting at B21) down to H20 in the copy file
    for i in range(21, 21 + column_length):  # Loop through rows in column B in the upload file
        value_from_B = xls_sheet.cell_value(i - 1, 1)  # B21 is (row 20, col 1) in zero-based indexing
        source_ws[f'H{i - 1}'] = value_from_B  # Paste into H20 and down
        print(f"Pasting {value_from_B} from B{i} to H{i - 1}")  # Debugging print

    # Now copy from upload file column C (starting at C21) down to I20 in the copy file
    for i in range(21, 21 + column_length):  # Loop through rows in column C in the upload file
        value_from_C = xls_sheet.cell_value(i - 1, 2)  # C21 is (row 20, col 2) in zero-based indexing
        source_ws[f'I{i - 1}'] = value_from_C  # Paste into I20 and down
        print(f"Pasting {value_from_C} from C{i} to I{i - 1}")  # Debugging print

    # Now copy from upload file column E (starting at E21) down to L20 in the copy file
    for i in range(21, 21 + column_length):  # Loop through rows in column E in the upload file
        value_from_E = xls_sheet.cell_value(i - 1, 4)  # E21 is (row 20, col 4) in zero-based indexing
        source_ws[f'L{i - 1}'] = value_from_E  # Paste into L20 and down
        print(f"Pasting {value_from_E} from E{i} to L{i - 1}")  # Debugging print

    # Save the updated copy
    source_wb.save(dest_file)

def process_ChewyASN(file_path):
    """Main function to process Chewy ASN files."""
    current_date = datetime.datetime.now().strftime("%m.%d.%Y")

    # Determine if the file is XLSX or XLS and extract the PO number
    if file_path.endswith('.xlsx'):
        uploaded_wb = load_workbook(file_path)
        po_number = uploaded_wb.active['C4'].value
    elif file_path.endswith('.xls'):
        xls_book = xlrd.open_workbook(file_path)
        po_number = xls_book.sheet_by_index(0).cell_value(3, 2)

    # Define the backup file path in FINISHED_FOLDER with a Chewy subfolder
    backup_file = os.path.join(FINISHED_FOLDER, f"Chewy/Chewy 856 ASN PO {po_number} {current_date}.xlsx")
    
    # Ensure the directory exists
    os.makedirs(os.path.dirname(backup_file), exist_ok=True)

    # Perform the copy or conversion based on file type
    if file_path.endswith('.xlsx'):
        copy_xlsx_data(file_path, backup_file)
    elif file_path.endswith('.xls'):
        convert_xls_data(file_path, backup_file)

    return backup_file, po_number
