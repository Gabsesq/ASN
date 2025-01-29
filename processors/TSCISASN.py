from openpyxl import load_workbook
import xlrd
import datetime
import os
import sys
import tempfile
from ExcelHelpers import (
    QTY_total, resource_path, oneToMany, manyToMany, get_current_date, extract_po_number, format_cells_as_text, align_cells_left, get_column_length, FINISHED_FOLDER
)



# Define source file for TSC IS
source_asn_xls = resource_path("assets/TSC IS/Master Template Tractor Supply IS ASN.xlsx")

# Check if the file exists
if not os.path.exists(source_asn_xls):
    print("File does not exist at:", source_asn_xls)
else:
    print("File found at:", source_asn_xls)


# Function to copy data from uploaded .xlsx file to the ASN .xlsx backup
def copy_xlsx_data(uploaded_file, dest_file):
    try:
        # For xlsx files, we'll convert the template first
        output_wb = load_workbook(source_asn_xls.replace('.xls', '.xlsx'))
        uploaded_wb = load_workbook(uploaded_file)
        
        output_ws = output_wb.active
        uploaded_ws = uploaded_wb.active

        format_cells_as_text(output_ws)
        align_cells_left(output_ws)

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
                output_ws[copy_cell] = value
                print(f"Copied value '{value}' from {upload_cell} to {copy_cell}")
            except Exception as e:
                print(f"Error copying from {upload_cell} to {copy_cell}: {str(e)}")

        # Calculate dynamic column length starting from A18
        column_length = get_column_length(uploaded_ws, start_row=18)

        # Copy value from 'C4' into 'B17' and down using oneToMany
        oneToMany(uploaded_ws, output_ws, 3, 2, 'B', 17, column_length)

        format_cells_as_text(output_ws)
        align_cells_left(output_ws)
        align_cells_left(output_ws)
        output_wb.save(dest_file)
        print(f"Saved file successfully as {dest_file}")
        return True

    except Exception as e:
        print(f"Error in copy_xlsx_data: {str(e)}")
        return False

# Function to convert and copy data from .xls files
def convert_xls_data(uploaded_file, dest_file):
    try:
        # Read the source template using xlrd
        source_book = xlrd.open_workbook(source_asn_xls)
        source_sheet = source_book.sheet_by_index(0)
        
        # Read the uploaded file using xlrd
        xls_book = xlrd.open_workbook(uploaded_file)
        xls_sheet = xls_book.sheet_by_index(0)

        # Create a new workbook for output
        output_wb = load_workbook(source_asn_xls.replace('.xls', '.xlsx'))
        output_ws = output_wb.active

        format_cells_as_text(output_ws)
        align_cells_left(output_ws)

        # Mapping uploaded cells to copy cells with error handling
        data_map = {
            (13, 1): 'E3', (13, 3): 'E4', (13, 4): 'E5',
            (13, 9): 'E6', (13, 10): 'E7', (13, 11): 'E8',
            (8, 2): 'B11', (6,2) : 'E11'
        }
        
        for (row, col), copy_cell in data_map.items():
            try:
                value = xls_sheet.cell_value(row, col)
                output_ws[copy_cell] = value
                print(f"Copied value '{value}' from ({row}, {col}) to {copy_cell}")
            except IndexError as e:
                print(f"IndexError: ({row}, {col}) is out of range: {str(e)}")
            except Exception as e:
                print(f"Unexpected error copying ({row}, {col}) to {copy_cell}: {str(e)}")

        # Calculate dynamic column length starting from A17
        column_length = get_column_length(xls_sheet, start_row=17)

        # Perform many-to-many copy operations with debugging
        manyToMany(xls_sheet, output_ws, 18, 0, 'A', 17, column_length)  # Item No
        oneToMany(xls_sheet, output_ws, 3, 2, 'B', 17, column_length-1)  # Copy 'C4' to 'B17'
        manyToMany(xls_sheet, output_ws, 18, 4, 'D', 17, column_length)  # UPS
        manyToMany(xls_sheet, output_ws, 18, 5, 'E', 17, column_length)  # Buyer Part
        manyToMany(xls_sheet, output_ws, 18, 6, 'F', 17, column_length)  # Vendor Part
        manyToMany(xls_sheet, output_ws, 18, 1, 'G', 17, column_length)  # QTY
        manyToMany(xls_sheet, output_ws, 18, 2, 'H', 17, column_length)  # UOM
        manyToMany(xls_sheet, output_ws, 18, 8, 'I', 17, column_length)  # Description

        # Sum the QTY values and place the total
        qty_total = QTY_total(xls_sheet, 18, 1)
        output_ws['E13'] = qty_total
        output_ws['B13'] = qty_total
        print(f"Total QTY placed in E13 and B13: {qty_total}")

        output_wb.save(dest_file)
        print(f"Saved file successfully as {dest_file}")
        return True

    except Exception as e:
        print(f"Error in convert_xls_data: {str(e)}")
        return False


def process_TSCIS(file_path):
    """Main function to process TSC files."""
    try:
        current_date = get_current_date()
        po_number = extract_po_number(file_path, is_xlsx=file_path.endswith('.xlsx'))
        
        # Fix the path separator issue by using os.path.join
        backup_file = os.path.join(FINISHED_FOLDER, 'TSCIS', f"Tractor Supply IS ASN {po_number} {current_date}.xlsx")
        
        # Ensure directory exists
        os.makedirs(os.path.dirname(backup_file), exist_ok=True)

        success = False
        if file_path.endswith('.xlsx'):
            success = copy_xlsx_data(file_path, backup_file)
        elif file_path.endswith('.xls'):
            success = convert_xls_data(file_path, backup_file)
        
        if not success:
            raise Exception("Failed to process file")
            
        if not os.path.exists(backup_file):
            raise Exception(f"Output file was not created at {backup_file}")
            
        return backup_file, po_number

    except Exception as e:
        print(f"Error in process_TSC: {str(e)}")
        raise  # Re-raise the exception to be caught by the main app


