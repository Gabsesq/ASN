from openpyxl import load_workbook
import xlrd
import os
from ExcelHelpers import (
    oneToMany, manyToMany, get_current_date, extract_po_number, 
    format_cells_as_text, align_cells_left, get_column_length, QTY_total
)

# Use /tmp for Vercel, or 'tmp/' for local development
TMP_PATH = '/tmp' if os.environ.get('VERCEL') else 'tmp'

# Ensure the tmp/ directory exists in local development
if not os.path.exists(TMP_PATH):
    os.makedirs(TMP_PATH)

# Define source template file for TSC
source_asn_xlsx = "assets/TSC/Blank TSC ASN.xlsx"

def copy_xlsx_data(uploaded_file, dest_file):
    """Copy data from an uploaded .xlsx file to the ASN backup."""
    try:
        uploaded_wb = load_workbook(uploaded_file)
        source_wb = load_workbook(source_asn_xlsx)
        uploaded_ws = uploaded_wb.active
        source_ws = source_wb.active

        format_cells_as_text(source_ws)
        align_cells_left(source_ws)

        # Data mapping from uploaded to template
        data_map = {
            'B14': 'E3', 'D14': 'E4', 'E14': 'E5',
            'J14': 'E6', 'K14': 'E7', 'L14': 'E8',
            'C9': 'B11'
        }

        # Transfer data
        for upload_cell, copy_cell in data_map.items():
            value = uploaded_ws[upload_cell].value
            source_ws[copy_cell] = value
            print(f"Copied value '{value}' from {upload_cell} to {copy_cell}")

        # Perform other operations
        column_length = get_column_length(uploaded_ws, start_row=18)
        oneToMany(uploaded_ws, source_ws, 3, 2, 'B', 17, column_length)

        # Save the file
        source_wb.save(dest_file)
        print(f"File saved successfully as {dest_file}")

    except Exception as e:
        print(f"Error in copy_xlsx_data: {str(e)}")

def convert_xls_data(uploaded_file, dest_file):
    """Convert and copy data from .xls to .xlsx."""
    try:
        xls_book = xlrd.open_workbook(uploaded_file)
        source_wb = load_workbook(source_asn_xlsx)
        source_ws = source_wb.active
        xls_sheet = xls_book.sheet_by_index(0)

        format_cells_as_text(source_ws)
        align_cells_left(source_ws)

        # Data mapping
        data_map = {
            (13, 1): 'E3', (13, 3): 'E4', (13, 4): 'E5',
            (13, 9): 'E6', (13, 10): 'E7', (13, 11): 'E8',
            (8, 2): 'B11'
        }

        # Transfer data
        for (row, col), copy_cell in data_map.items():
            value = xls_sheet.cell_value(row, col)
            source_ws[copy_cell] = value

        # Perform operations
        column_length = get_column_length(xls_sheet, start_row=17)
        manyToMany(xls_sheet, source_ws, 18, 0, 'A', 17, column_length)

        # Save the file
        source_wb.save(dest_file)
        print(f"File saved successfully as {dest_file}")

    except Exception as e:
        print(f"Error in convert_xls_data: {str(e)}")

def process_TSC(file_path):
    """Process TSC files and save the output to TMP_PATH."""
    try:
        current_date = get_current_date()
        po_number = extract_po_number(file_path, file_path.endswith('.xlsx'))

        # Construct the output path in TMP_PATH
        output_file = f"{TMP_PATH}/TSC/Tractor Supply ASN {po_number} {current_date}.xlsx"

        # Ensure the directory exists
        os.makedirs(os.path.dirname(output_file), exist_ok=True)

        # Process the file based on type
        if file_path.endswith('.xlsx'):
            copy_xlsx_data(file_path, output_file)
        elif file_path.endswith('.xls'):
            convert_xls_data(file_path, output_file)

        return output_file

    except Exception as e:
        print(f"Error in process_TSC: {str(e)}")
        return None
