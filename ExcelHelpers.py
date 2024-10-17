import os
import datetime
from openpyxl.styles import NamedStyle, Alignment


def format_cells_as_text(worksheet):
    """Format only the cells with data as text."""
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value is not None:  # Only apply formatting if there's data
                cell.number_format = "@"

def align_cells_left(worksheet):
    """Align only the cells with data to the left."""
    left_alignment = Alignment(horizontal='left')
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value is not None:  # Only align cells with data
                cell.alignment = left_alignment

# Helper function to get the current date as a string
def get_current_date():
    return datetime.datetime.now().strftime("%m.%d.%Y")

# Helper function to create a folder if it doesn't exist
def create_folder(folder_path):
    os.makedirs(folder_path, exist_ok=True)

# Helper function to extract PO number from XLSX or XLS files
def extract_po_number(file_path, is_xlsx=True):
    if is_xlsx:
        from openpyxl import load_workbook
        wb = load_workbook(file_path)
        return wb.active['C4'].value
    else:
        import xlrd
        xls_book = xlrd.open_workbook(file_path)
        return xls_book.sheet_by_index(0).cell_value(3, 2)