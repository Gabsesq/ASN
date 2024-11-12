from openpyxl import load_workbook
import xlrd
import os
from ExcelHelpers import (
    resource_path, FINISHED_FOLDER, format_cells_as_text, align_cells_left, manyToMany, oneToMany,
    typedValue, QTY_total, get_column_length, create_folder, extract_po_number, get_current_date, generate_rows
)

# Define source ASN template file path
source_asn_xlsx = resource_path("assets/Thrive Market/Thrive Market 856 Master Template.xlsx")



def copy_xlsx_data(uploaded_file, dest_file):
    """Copy data from an uploaded .xlsx file to the ASN backup."""
    uploaded_wb = load_workbook(uploaded_file)
    source_wb = load_workbook(source_asn_xlsx)
    uploaded_ws = uploaded_wb.active
    source_ws = source_wb.active

    format_cells_as_text(source_ws)
    align_cells_left(source_ws)


    # Copy data dynamically using helper functions
    manyToMany(uploaded_ws, source_ws, 17, 0, 'A', 19, get_column_length(uploaded_ws, 17))
    oneToMany(uploaded_ws, source_ws, 3, 2, 'B', 19, get_column_length(uploaded_ws, 17))
    oneToMany(uploaded_ws, source_ws, 3, 7, 'C', 19, get_column_length(uploaded_ws, 17))

    # Use a static value for a specific column
    typedValue(source_ws, "Fed Ex", 'D', 19, get_column_length(uploaded_ws, 17))

    # Calculate total quantity and place it in 'F14'
    total_qty = QTY_total(uploaded_ws, 17, 1)
    source_ws['F14'] = total_qty

    # Save the updated workbook
    source_wb.save(dest_file)
    print(f"File saved successfully as {dest_file}.")

def convert_xls_data(uploaded_file, dest_file):
    """Convert and copy data from .xls to .xlsx backup."""
    xls_book = xlrd.open_workbook(uploaded_file)
    source_wb = load_workbook(source_asn_xlsx)
    source_ws = source_wb.active
    xls_sheet = xls_book.sheet_by_index(0)

    format_cells_as_text(source_ws)
    align_cells_left(source_ws)

    # Define your data maps
    data_map = {
        (12, 1): 'G3', (12, 3): 'G4', (12, 4): 'G5',
        (12, 7): 'G6', (12, 9): 'G7', (12, 10): 'G8', (12, 11): 'G9',
    }


    for (row, col), copy_cell in data_map.items():
        value = xls_sheet.cell_value(row, col)
        source_ws[copy_cell] = value

    # Copy data dynamically using helper functions
    manyToMany(xls_sheet, source_ws, 17, 0, 'A', 19, get_column_length(xls_sheet, 17)) # Line number
    oneToMany(xls_sheet, source_ws, 3, 2, 'B', 19, get_column_length(xls_sheet, 17)) # PO
    oneToMany(xls_sheet, source_ws, 3, 7, 'C', 19, get_column_length(xls_sheet, 17)) # PO date
    manyToMany(xls_sheet, source_ws, 17, 6, 'E', 19, get_column_length(xls_sheet, 17)) # vendor part
    manyToMany(xls_sheet, source_ws, 17, 5, 'F', 19, get_column_length(xls_sheet, 17)) # UPC
    manyToMany(xls_sheet, source_ws, 17, 1, 'G', 19, get_column_length(xls_sheet, 17)) # QTY NEED TO CHANGE
    manyToMany(xls_sheet, source_ws, 17, 2, 'H', 19, get_column_length(xls_sheet, 17)) # UOM
    manyToMany(xls_sheet, source_ws, 17, 7, 'I', 19, get_column_length(xls_sheet, 17)) # pack
    manyToMany(xls_sheet, source_ws, 17, 4, 'J', 19, get_column_length(xls_sheet, 17)) # Description

 # Generate rows based on quantity and write them into the destination sheet
    rows_to_write = generate_rows(xls_sheet, start_row=17, qty_column=6, column_count=xls_sheet.ncols)

    print(f"Total rows generated: {len(rows_to_write)}")

    # Write the generated rows into the destination worksheet
    for i, row_data in enumerate(rows_to_write):
        for j, value in enumerate(row_data):
            source_ws.cell(row=19 + i, column=j + 1, value=value)
            print(f"Wrote value '{value}' to cell ({19 + i}, {j + 1})")


    total_qty = QTY_total(xls_sheet, 17, 1)
    source_ws['F15'] = total_qty

    format_cells_as_text(source_ws)
    align_cells_left(source_ws)
    align_cells_left(source_ws)

    # Save the updated workbook
    source_wb.save(dest_file)
    print(f"File saved successfully as {dest_file}.")

def process_ThriveASN(file_path):
    """Main function to process Thrive ASN files."""
    current_date = get_current_date()
    po_number = extract_po_number(file_path, file_path.endswith('.xlsx'))

    # Define the destination backup file path
    backup_file = os.path.join(FINISHED_FOLDER, f"Thrive/Thrive ASN PO {po_number} {current_date}.xlsx")
    
    # Ensure the Thrive folder exists in the Finished directory
    os.makedirs(os.path.dirname(backup_file), exist_ok=True)
    print(f"Resolved backup file path: {backup_file}")  # Debug line to confirm path

    # Process based on file type
    if file_path.endswith('.xlsx'):
        copy_xlsx_data(file_path, backup_file)
    elif file_path.endswith('.xls'):
        convert_xls_data(file_path, backup_file)

    return backup_file, po_number
