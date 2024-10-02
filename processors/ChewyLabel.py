from openpyxl import load_workbook
import xlrd
import datetime

# Define source files and destination copies for Chewy
source_label_xlsx = "assets\Chewy\Chewy UCC128 Label Request - Copy.xlsx"

# Function to copy data from uploaded .xlsx file to specific cells in the ASN .xlsx backup
def copy_xlsx_data(uploaded_file, dest_file):
    uploaded_wb = load_workbook(uploaded_file)
    source_wb = load_workbook(source_label_xlsx)
    uploaded_ws = uploaded_wb.active
    source_ws = source_wb.active

    # Mapping uploaded cells to copy cells
    data_map = {
        'B16': 'F3', 'E16': 'F4', 'H16': 'F5',
        'J16': 'F6', 'K16': 'F7', 'L16': 'F8'
    }

    # Transfer data from the uploaded file to the backup copy
    for upload_cell, copy_cell in data_map.items():
        source_ws[copy_cell] = uploaded_ws[upload_cell].value

    # Track numbers from A21 and below, copy to A20 in the copy
    row = 15
    data_to_copy = []
    while uploaded_ws[f'A{row}'].value is not None:
        data_to_copy.append(uploaded_ws[f'A{row}'].value)
        row += 1

    for i, value in enumerate(data_to_copy):
        source_ws[f'A{14 + i}'] = value

    count = len(data_to_copy)
    value_from_upload = uploaded_ws['C4'].value

    if value_from_upload is not None:
        for i in range(14, 14 + count):
            source_ws[f'A{i}'] = value_from_upload

    source_wb.save(dest_file)

# Function to convert .xls to .xlsx and transfer data to a backup of ASN copy
def convert_xls_data(uploaded_file, dest_file):
    xls_book = xlrd.open_workbook(uploaded_file)
    source_wb = load_workbook(source_label_xlsx)
    source_ws = source_wb.active
    xls_sheet = xls_book.sheet_by_index(0)

    # Mapping uploaded cells to copy cells
    data_map = {
        (15, 1): 'F3', (15, 4): 'F4', (15, 7): 'F5',
        (15, 9): 'F6', (15, 10): 'F7', (15, 11): 'F8'
    }

    for (row, col), copy_cell in data_map.items():
        value = xls_sheet.cell_value(row, col)
        source_ws[copy_cell] = value

    # Track numbers from A21 and below, copy to A20 in the copy
    row = 15
    copy_row = 14
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
    for i in range(15, 15 + column_length):  # Loop through rows in column H in the upload file
        value_from_H = xls_sheet.cell_value(i - 1, 7)  # H21 is (row 20, col 7) in zero-based indexing
        source_ws[f'E{i - 1}'] = value_from_H  # Paste into E20 and down
        print(f"Pasting {value_from_H} from H{i} to E{i - 1}")  # Debugging print

    # Now copy value from C4 into B column (B20 and down for column_length rows)
    value_from_upload_C4 = xls_sheet.cell_value(3, 2)
    if value_from_upload_C4 is not None:
        for i in range(14, 14 + column_length):  # Copy value to B20 through B (dynamic based on column A length)
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

# Main function to process Chewy files
def process_label(file_path):
    current_date = datetime.datetime.now().strftime("%m.%d.%Y")

    if file_path.endswith('.xlsx'):
        uploaded_wb = load_workbook(file_path)
        po_number = uploaded_wb.active['C4'].value
        backup_file = f"Finished/Chewy/Chewy UCC128 Label Request PO {po_number} {current_date}.xlsx"
        copy_xlsx_data(file_path, backup_file)

    elif file_path.endswith('.xls'):
        xls_book = xlrd.open_workbook(file_path)
        po_number = xls_book.sheet_by_index(0).cell_value(3, 2)
        backup_file = f"Finished/Chewy/Chewy UCC128 Label Request PO {po_number} {current_date}.xlsx"
        convert_xls_data(file_path, backup_file)
