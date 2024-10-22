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
    
def manyToMany(xls_sheet, source_ws, start_row, start_col, dest_col, dest_start_row, column_length):
    """
    Copies values from the source sheet to the destination sheet, row-by-row.
    Handles cases where only a single row exists.
    """
    print(f"Copying from column {start_col} (starting at row {start_row}) "
          f"to {dest_col}{dest_start_row} for {column_length} rows.")

    for i in range(column_length):
        try:
            value = xls_sheet.cell_value(start_row + i - 1, start_col)
            source_ws[f'{dest_col}{dest_start_row + i}'] = value
            print(f"Pasted '{value}' to {dest_col}{dest_start_row + i}")
        except IndexError as e:
            print(f"IndexError at row {start_row + i - 1}, col {start_col}: {str(e)}")
        except Exception as e:
            print(f"Unexpected error: {str(e)}")
# ExcelHelpers.py

def oneToMany(xls_sheet, source_ws, row, col, target_column, start_row, column_length):
    """
    Copies a value from one specific cell and pastes it into multiple rows in a target column.

    Parameters:
        xls_sheet: The uploaded XLS sheet object.
        source_ws: The destination worksheet object.
        row (int): Zero-based row index of the source cell.
        col (int): Zero-based column index of the source cell.
        target_column (str): The column letter in the destination sheet (e.g., 'E').
        start_row (int): The starting row in the destination sheet (e.g., 8).
        column_length (int): Number of rows to paste the value into.
    """
    try:
        # Fetch the value from the specific cell
        value = xls_sheet.cell_value(row, col)
        print(f"Copying value '{value}' from ({row + 1}, {col + 1})")

        # Paste the value into the specified column for the given length
        for i in range(start_row, start_row + column_length):
            source_ws[f'{target_column}{i}'] = value
            print(f"Pasting '{value}' into {target_column}{i}")

    except IndexError as e:
        print(f"Error accessing cell ({row}, {col}): {str(e)}")
    except Exception as e:
        print(f"Unexpected error: {str(e)}")

def typedValue(source_ws, static_value, target_column, start_row, column_length):
    """
    Pastes a static value (like "N/A") into multiple rows in a target column.

    Parameters:
        source_ws: The destination worksheet object.
        static_value: The value to paste (e.g., "N/A").
        target_column (str): The column letter in the destination sheet (e.g., 'D').
        start_row (int): The starting row in the destination sheet (e.g., 19).
        column_length (int): Number of rows to paste the value into.
    """
    try:
        print(f"Using static value '{static_value}'")

        # Paste the static value into the specified column for the given length
        for i in range(start_row, start_row + column_length):
            source_ws[f'{target_column}{i}'] = static_value
            print(f"Pasting '{static_value}' into {target_column}{i}")

    except Exception as e:
        print(f"Unexpected error: {str(e)}")

def get_column_length(sheet, start_row):
    """Calculate the number of non-empty rows, ensuring at least one row is processed."""
    column_length = 0
    while True:
        try:
            value = sheet.cell_value(start_row - 1, 0)  # Column A (index 0)
            print(f"Row {start_row}: Value in A = '{value}'")
            if value:
                column_length += 1
                start_row += 1
            else:
                break
        except IndexError as e:
            print(f"IndexError accessing row {start_row - 1}, column 0: {str(e)}")
            break

    # Ensure at least one row is processed, even if only one row exists
    if column_length == 0:
        column_length = 1
        print("Column length adjusted to 1 to handle single-row data.")

    print(f"Final Column Length: {column_length}")
    return column_length