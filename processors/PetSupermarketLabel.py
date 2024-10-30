import os
import datetime
from openpyxl import load_workbook
import xlrd
from ExcelHelpers import (
    format_cells_as_text, align_cells_left
)

# Define source files and destination copies for Pet Supermarket
source_label_xlsx = "assets\Pet Supermarket\Blank Pet Supermarket UCC 128 Label Request.xlsx"

# Function to copy data from uploaded .xlsx file to specific cells in the label request .xlsx backup
def copy_xlsx_data(uploaded_file, dest_file):
    uploaded_wb = load_workbook(uploaded_file)
    source_wb = load_workbook(source_label_xlsx)
    uploaded_ws = uploaded_wb.active
    source_ws = source_wb.active

    # Mapping uploaded cells to copy cells
    data_map = {
        'C4': 'A14',  # Copy C4 to A14
        'L12': 'B14',  # Copy L12 to B14
    }

    # Transfer data from the uploaded file to the backup copy
    for upload_cell, copy_cell in data_map.items():
        source_ws[copy_cell] = uploaded_ws[upload_cell].value

    # Set static values
    source_ws['C14'] = "SAIA"
    source_ws['E14'] = "mixed"
    source_ws['G14'] = "mixed"
    source_ws['H14'] = 1

    align_cells_left(source_ws)
    # Save the updated copy
    source_wb.save(dest_file)

# Function to convert .xls to .xlsx and transfer data to a backup of the label request copy
def convert_xls_data(uploaded_file, dest_file):
    xls_book = xlrd.open_workbook(uploaded_file)
    source_wb = load_workbook(source_label_xlsx)
    source_ws = source_wb.active
    xls_sheet = xls_book.sheet_by_index(0)
    align_cells_left(source_ws)

    # Mapping uploaded cells to copy cells
    data_map = {
        (3, 2): 'A14',  # C4 is (row 3, col 2) in zero-based indexing, copying to A14
        (11, 11): 'B14'  # L12 is (row 11, col 11) in zero-based indexing, copying to B14
    }

    # Transfer data from the uploaded .xls file to the backup copy
    for (row, col), copy_cell in data_map.items():
        value = xls_sheet.cell_value(row, col)
        source_ws[copy_cell] = value

    # Set static values
    source_ws['C14'] = "SAIA"
    source_ws['E14'] = "mixed"
    source_ws['G14'] = "mixed"
    source_ws['H14'] = 1

    align_cells_left(source_ws)
    # Save the updated copy
    source_wb.save(dest_file)

# Main function to process Pet Supermarket Label files based on file type
def process_PetSuperLabel(file_path):
    current_date = datetime.datetime.now().strftime("%m.%d.%Y")

    # Ensure the 'Finished/PetSupermarket' directory exists
    finished_folder = "Finished/PetSupermarket"
    if not os.path.exists(finished_folder):
        os.makedirs(finished_folder)

    # Process .xlsx file
    if file_path.endswith('.xlsx'):
        uploaded_wb = load_workbook(file_path)
        po_number = uploaded_wb.active['C4'].value
        backup_file = f"{finished_folder}/Pet Supermarket Label Request PO {po_number} {current_date}.xlsx"
        copy_xlsx_data(file_path, backup_file)

    # Process .xls file
    elif file_path.endswith('.xls'):
        xls_book = xlrd.open_workbook(file_path)
        po_number = xls_book.sheet_by_index(0).cell_value(3, 2)
        backup_file = f"{finished_folder}/Pet Supermarket Label Request PO {po_number} {current_date}.xlsx"
        convert_xls_data(file_path, backup_file)
