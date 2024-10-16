from openpyxl import load_workbook
import xlrd
import datetime

# Define source files and destination copies for Chewy
source_asn_xlsx = "assets/Thrive Market/New Blank Thrive ASN.xlsx"

# Function to copy data from uploaded .xlsx file to specific cells in the ASN .xlsx backup
def copy_xlsx_data(uploaded_file, dest_file):
    uploaded_wb = load_workbook(uploaded_file)
    source_wb = load_workbook(source_asn_xlsx)
    uploaded_ws = uploaded_wb.active
    source_ws = source_wb.active

     # Mapping uploaded cells to copy cells
    data_map = {
        (15, 1): 'E3', (15, 3): 'E4', (15, 4): 'E5',
        (15, 9): 'E6', (15, 10): 'E7', (15, 11): 'E8',
        (7, 2): 'E14'
    }
    # Transfer data from the uploaded file to the backup copy
    for upload_cell, copy_cell in data_map.items():
        source_ws[copy_cell] = uploaded_ws[upload_cell].value
    
    # Track data from A17 and below, copy to A19 and below in the destination
    row = 17
    data_to_copy = []
    
    # Collect data from A17 downwards until a blank cell is found
    while uploaded_ws[f'A{row}'].value is not None:
        data_to_copy.append(uploaded_ws[f'A{row}'].value)
        row += 1

    # Paste the collected data into A19 downwards in the destination file
    for i, value in enumerate(data_to_copy):
        source_ws[f'A{19 + i}'] = value

    # Count of rows copied
    count = len(data_to_copy)
    
    # Copy value from C4 in the uploaded file into B19 downwards (same length as column A data)
    value_from_C4 = uploaded_ws['C4'].value
    if value_from_C4 is not None:
        for i in range(count):
            source_ws[f'B{19 + i}'] = value_from_C4

    # Copy value from H4 in the uploaded file into C19 downwards (same length as column A data)
    value_from_H4 = uploaded_ws['H4'].value
    if value_from_H4 is not None:
        for i in range(count):
            source_ws[f'C{19 + i}'] = value_from_H4

    # Now copy data from G17 downwards into E19 downwards in the destination file
    row = 17
    data_to_copy = []
    while uploaded_ws[f'G{row}'].value is not None:
        data_to_copy.append(uploaded_ws[f'G{row}'].value)
        row += 1
    
    for i, value in enumerate(data_to_copy):
        source_ws[f'E{19 + i}'] = value

    # Copy data from F17 downwards into F19 downwards in the destination file
    row = 17
    data_to_copy = []
    while uploaded_ws[f'F{row}'].value is not None:
        data_to_copy.append(uploaded_ws[f'F{row}'].value)
        row += 1
    
    for i, value in enumerate(data_to_copy):
        source_ws[f'F{19 + i}'] = value

    # Copy data from B17 downwards into G19 downwards in the destination file
    row = 17
    data_to_copy = []
    while uploaded_ws[f'B{row}'].value is not None:
        data_to_copy.append(uploaded_ws[f'B{row}'].value)
        row += 1
    
    for i, value in enumerate(data_to_copy):
        source_ws[f'G{19 + i}'] = value

    # Copy data from C17 downwards into H19 downwards in the destination file
    row = 17
    data_to_copy = []
    while uploaded_ws[f'C{row}'].value is not None:
        data_to_copy.append(uploaded_ws[f'C{row}'].value)
        row += 1
    
    for i, value in enumerate(data_to_copy):
        source_ws[f'H{19 + i}'] = value

    # Copy data from H17 downwards into I19 downwards in the destination file
    row = 17
    data_to_copy = []
    while uploaded_ws[f'H{row}'].value is not None:
        data_to_copy.append(uploaded_ws[f'H{row}'].value)
        row += 1
    
    for i, value in enumerate(data_to_copy):
        source_ws[f'I{19 + i}'] = value

    # Copy data from E17 downwards into J19 downwards in the destination file
    row = 17
    data_to_copy = []
    while uploaded_ws[f'E{row}'].value is not None:
        data_to_copy.append(uploaded_ws[f'E{row}'].value)
        row += 1
    
    for i, value in enumerate(data_to_copy):
        source_ws[f'J{19 + i}'] = value

    # Save the updated copy
    source_wb.save(dest_file)

# Function to convert .xls to .xlsx and transfer data to a backup of ASN copy
def convert_xls_data(uploaded_file, dest_file):
    xls_book = xlrd.open_workbook(uploaded_file)
    source_wb = load_workbook(source_asn_xlsx)
    source_ws = source_wb.active
    xls_sheet = xls_book.sheet_by_index(0)

    # Map specific cells from the uploaded .xls to the .xlsx backup
    data_map = {
        (11, 1): 'E4',   # B12 in .xlsx corresponds to row 11 (0-based index) and column 1 (B)
        (11, 3): 'E5',   # D12 in .xlsx corresponds to row 11 and column 3 (D)
        (11, 4): 'E6',   # E12 in .xlsx corresponds to row 11 and column 4 (E)
        (11, 9): 'E7',   # J12 in .xlsx corresponds to row 11 and column 9 (J)
        (11, 10): 'E8',  # K12 in .xlsx corresponds to row 11 and column 10 (K)
        (11, 11): 'E9'   # L12 in .xlsx corresponds to row 11 and column 11 (L)
    }

    # Copy the mapped cells from the .xls file to the .xlsx file
    for (row, col), copy_cell in data_map.items():
        value = xls_sheet.cell_value(row, col)
        source_ws[copy_cell] = value

    # Track data from A17 and below in .xls, copy to A19 in the .xlsx destination file
    row = 17
    data_to_copy = []
    
    while True:
        try:
            value = xls_sheet.cell_value(row - 1, 0)  # Column A is index 0 (zero-based)
            if value:
                data_to_copy.append(value)
                row += 1
            else:
                break
        except IndexError:
            break

    # Paste data to A19 in the destination file
    for i, value in enumerate(data_to_copy):
        source_ws[f'A{19 + i}'] = value

    # Get the number of rows copied
    count = len(data_to_copy)

    # Copy value from C4 in the .xls file to B19 and down in the .xlsx file
    value_from_C4 = xls_sheet.cell_value(3, 2)
    if value_from_C4 is not None:
        for i in range(count):
            source_ws[f'B{19 + i}'] = value_from_C4

    # Copy value from H4 in the .xls file to C19 and down in the .xlsx file
    value_from_H4 = xls_sheet.cell_value(3, 7)
    if value_from_H4 is not None:
        for i in range(count):
            source_ws[f'C{19 + i}'] = value_from_H4

    # Copy data from G17 down in the .xls file to E19 and down in the .xlsx file
    row = 17
    while True:
        try:
            value = xls_sheet.cell_value(row - 1, 6)  # Column G is index 6 (zero-based)
            if value:
                source_ws[f'E{19 + row - 17}'] = value
                row += 1
            else:
                break
        except IndexError:
            break

    # Copy data from F17 down in the .xls file to F19 and down in the .xlsx file
    row = 17
    while True:
        try:
            value = xls_sheet.cell_value(row - 1, 5)  # Column F is index 5 (zero-based)
            if value:
                source_ws[f'F{19 + row - 17}'] = value
                row += 1
            else:
                break
        except IndexError:
            break

    # Copy data from B17 down in the .xls file to G19 and down in the .xlsx file
    row = 17
    while True:
        try:
            value = xls_sheet.cell_value(row - 1, 1)  # Column B is index 1 (zero-based)
            if value:
                source_ws[f'G{19 + row - 17}'] = value
                row += 1
            else:
                break
        except IndexError:
            break

    # Copy data from C17 down in the .xls file to H19 and down in the .xlsx file
    row = 17
    while True:
        try:
            value = xls_sheet.cell_value(row - 1, 2)  # Column C is index 2 (zero-based)
            if value:
                source_ws[f'H{19 + row - 17}'] = value
                row += 1
            else:
                break
        except IndexError:
            break

    # Copy data from H17 down in the .xls file to I19 and down in the .xlsx file
    row = 17
    while True:
        try:
            value = xls_sheet.cell_value(row - 1, 7)  # Column H is index 7 (zero-based)
            if value:
                source_ws[f'I{19 + row - 17}'] = value
                row += 1
            else:
                break
        except IndexError:
            break

    # Copy data from E17 down in the .xls file to J19 and down in the .xlsx file
    row = 17
    while True:
        try:
            value = xls_sheet.cell_value(row - 1, 4)  # Column E is index 4 (zero-based)
            if value:
                source_ws[f'J{19 + row - 17}'] = value
                row += 1
            else:
                break
        except IndexError:
            break

    # Save the updated copy
    source_wb.save(dest_file)


# Main function to process Chewy files
def process_ThriveASN(file_path):
    current_date = datetime.datetime.now().strftime("%m.%d.%Y")

    if file_path.endswith('.xlsx'):
        uploaded_wb = load_workbook(file_path)
        po_number = uploaded_wb.active['C4'].value
        backup_file = f"Finished/Thrive/Thrive ASN PO {po_number} {current_date}.xlsx"
        copy_xlsx_data(file_path, backup_file)

    elif file_path.endswith('.xls'):
        xls_book = xlrd.open_workbook(file_path)
        po_number = xls_book.sheet_by_index(0).cell_value(3, 2)
        backup_file = f"Finished/Thrive/Thrive ASN PO {po_number} {current_date}.xlsx"
        convert_xls_data(file_path, backup_file)
