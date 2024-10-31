from openpyxl import load_workbook
import xlrd
import datetime
from ExcelHelpers import (
    oneToMany, manyToMany, get_current_date, extract_po_number, format_cells_as_text, align_cells_left, get_column_length
)

# Define source file for TSC
source_asn_xlsx = "assets/TSC/Blank TSC ASN.xlsx"

# Function to copy data from uploaded .xlsx file to the ASN .xlsx backup
def copy_xlsx_data(uploaded_file, dest_file):
    try:
        uploaded_wb = load_workbook(uploaded_file)
        source_wb = load_workbook(source_asn_xlsx)
        uploaded_ws = uploaded_wb.active
        source_ws = source_wb.active

        format_cells_as_text(source_ws)
        align_cells_left(source_ws)

        # Mapping uploaded cells to copy cells
        data_map = {
            'B14': 'E3', 'D14': 'E4', 'E14': 'E5',
            'J14': 'E6', 'K14': 'E7', 'L14': 'E8',
            'C9': 'B11'
        }

        # Transfer data with debugging
        for upload_cell, copy_cell in data_map.items():
            try:
                value = uploaded_ws[upload_cell].value
                source_ws[copy_cell] = value
                print(f"Copied value '{value}' from {upload_cell} to {copy_cell}")
            except Exception as e:
                print(f"Error copying from {upload_cell} to {copy_cell}: {str(e)}")

        # Calculate dynamic column length starting from A18
        column_length = get_column_length(uploaded_ws, start_row=18)

        # Copy value from 'C4' into 'B17' and down using oneToMany
        oneToMany(uploaded_ws, source_ws, 3, 2, 'B', 17, column_length)

        source_wb.save(dest_file)
        print(f"Saved file successfully as {dest_file}")

    except Exception as e:
        print(f"Error in copy_xlsx_data: {str(e)}")

# Function to convert and copy data from .xls files
def convert_xls_data(uploaded_file, dest_file):
    try:
        xls_book = xlrd.open_workbook(uploaded_file)
        source_wb = load_workbook(source_asn_xlsx)
        source_ws = source_wb.active
        xls_sheet = xls_book.sheet_by_index(0)

        format_cells_as_text(source_ws)
        align_cells_left(source_ws)

        # Mapping uploaded cells to copy cells with error handling
        data_map = {
            (13, 1): 'E3', (13, 3): 'E4', (13, 4): 'E5',
            (13, 9): 'E6', (13, 10): 'E7', (13, 11): 'E8',
            (8, 2): 'B11'
        }
        for (row, col), copy_cell in data_map.items():
            try:
                value = xls_sheet.cell_value(row, col)
                source_ws[copy_cell] = value
                print(f"Copied value '{value}' from ({row}, {col}) to {copy_cell}")
            except IndexError as e:
                print(f"IndexError: ({row}, {col}) is out of range: {str(e)}")
            except Exception as e:
                print(f"Unexpected error copying ({row}, {col}) to {copy_cell}: {str(e)}")

        # Calculate dynamic column length starting from A17
        column_length = get_column_length(xls_sheet, start_row=17)

        # Perform many-to-many copy operations with debugging
        manyToMany(xls_sheet, source_ws, 18, 0, 'A', 17, column_length)  # Item No
        oneToMany(xls_sheet, source_ws, 3, 2, 'B', 17, column_length)    # Copy 'C4' to 'B17'
        manyToMany(xls_sheet, source_ws, 18, 4, 'D', 17, column_length)  # UPS
        manyToMany(xls_sheet, source_ws, 18, 5, 'E', 17, column_length)  # Buyer Part
        manyToMany(xls_sheet, source_ws, 18, 6, 'F', 17, column_length)  # Vendor Part
        manyToMany(xls_sheet, source_ws, 18, 1, 'G', 17, column_length)  # QTY
        manyToMany(xls_sheet, source_ws, 18, 2, 'H', 17, column_length)  # UOM
        manyToMany(xls_sheet, source_ws, 18, 8, 'I', 17, column_length)  # Description

        # Sum the QTY values from column B (index 1) and place the total in E13
        qty_total = sum_qty_values(xls_sheet, start_row=17, column_index=1, column_length=column_length)
        source_ws['E13'] = qty_total
        print(f"Total QTY placed in E13: {qty_total}")


        source_wb.save(dest_file)
        print(f"Saved file successfully as {dest_file}")

    except Exception as e:
        print(f"Error in convert_xls_data: {str(e)}")

def get_column_length(sheet, start_row):
    """Calculate the number of non-empty rows starting from a given row."""
    column_length = 0
    total_rows = sheet.nrows  # Ensure we don't exceed the number of rows

    while start_row <= total_rows:
        try:
            value = sheet.cell_value(start_row - 1, 0)  # Column A (index 0)
            print(f"Row {start_row}: Value in A = '{value}'")

            if value:
                column_length += 1
                start_row += 1
            else:
                break  # Stop when we encounter an empty cell
        except IndexError as e:
            print(f"IndexError accessing row {start_row - 1}, column 0: {str(e)}")
            break

    print(f"Final Column Length: {column_length}")
    return max(1, column_length)  # Ensure at least 1 row is counted


def sum_qty_values(sheet, start_row, column_index, column_length):
    """Sum the QTY values from the specified column and return the total."""
    qty_total = 0
    for i in range(start_row, start_row + column_length):
        try:
            value = sheet.cell_value(i - 1, column_index)
            print(f"Row {i}: Raw QTY value = '{value}'")

            # Handle text-formatted numbers by converting them to float
            if isinstance(value, str) and value.isnumeric():
                value = float(value)

            if isinstance(value, (int, float)):  # Ensure the value is numeric
                qty_total += value
            else:
                print(f"Skipping non-numeric value at row {i}: {value}")
        except IndexError as e:
            print(f"IndexError at row {i - 1}, column {column_index}: {str(e)}")
        except Exception as e:
            print(f"Unexpected error: {str(e)}")
    print(f"Final QTY Total: {qty_total}")
    return qty_total

def process_TSC(file_path):
    """Main function to process TSC files."""
    try:
        current_date = get_current_date()

        if file_path.endswith('.xlsx'):
            po_number = extract_po_number(file_path, is_xlsx=True)
            backup_file = f"Finished/TSC/Tractor Supply ASN {po_number} {current_date}.xlsx"
            copy_xlsx_data(file_path, backup_file)

        elif file_path.endswith('.xls'):
            po_number = extract_po_number(file_path, is_xlsx=False)
            backup_file = f"Finished/TSC/Tractor Supply ASN {po_number} {current_date}.xlsx"
            convert_xls_data(file_path, backup_file)
        
        # Return backup_file path to ensure it’s accessible for further processing
        return backup_file

    except Exception as e:
        print(f"Error in process_TSC: {str(e)}")
        return None  # Ensure None is returned if an error occurs

