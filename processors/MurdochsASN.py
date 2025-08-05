from openpyxl import load_workbook
import xlrd
import datetime
import os
from ExcelHelpers import (
    resource_path, FINISHED_FOLDER, format_cells_as_text, align_cells_left, manyToMany, oneToMany, typedValue
)
from upc_counts import counts  # Import the UPC counts dictionary

# Define source files and destination copies for Chewy
source_asn_xlsx = resource_path("assets/Murdochs/Blank Murdochs 856 ASN.xlsx")

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
        'B18': 'E3', 'D18': 'E4', 'E18': 'E5',
        'J18': 'E6', 'K18': 'E7', 'L18': 'E8',
        'C10': 'E14'
    }

    # Transfer data from the uploaded file to the backup copy
    for upload_cell, copy_cell in data_map.items():
        source_ws[copy_cell] = uploaded_ws[upload_cell].value

    # Manually assign values for cells that don't come from the uploaded file
    source_ws['B12'] = "FEDG"
    source_ws['B13'] = "Fedex"
    
    # Set current date in E11 and highlight it
    current_date = datetime.datetime.now().strftime("%m/%d/%Y")
    source_ws['E11'] = current_date
    
    # Highlight cell E11 with yellow background
    from openpyxl.styles import PatternFill
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    source_ws['E11'].fill = yellow_fill

    # Create carton-based lines from the uploaded Excel file
    carton_lines_count = create_carton_based_lines_from_xlsx(uploaded_ws, source_ws, start_row=23, target_start_row=19)
    
    print(f"Total carton lines created: {carton_lines_count}")
    
    # Additional copy and paste operations with helper functions
    oneToMany_xlsx(uploaded_ws, source_ws, row=4, col=2, target_column='B', start_row=19, column_length=carton_lines_count)  # PO
    oneToMany_xlsx(uploaded_ws, source_ws, row=4, col=7, target_column='C', start_row=19, column_length=carton_lines_count)  # PO Date

    typedValue(source_ws, static_value="NA", target_column='D', start_row=19, column_length=carton_lines_count)

    # Set total cartons in B15 and "C" in E13
    source_ws['B15'] = carton_lines_count
    source_ws['E13'] = "C"

    source_wb.save(dest_file)

def create_carton_based_lines_from_xlsx(uploaded_ws, source_ws, start_row=23, target_start_row=19):
    """
    Creates one line per carton instead of one line per product for XLSX files.
    For each product, calculates how many cartons are needed and creates that many lines.
    """
    carton_lines = []  # Store all the data for carton-based lines
    
    # Process each product row from the uploaded file
    row = start_row
    while uploaded_ws[f'A{row}'].value is not None:
        try:
            # Get product data from the uploaded file
            item_no = uploaded_ws[f'A{row}'].value  # Column A - Item number
            qty = int(uploaded_ws[f'B{row}'].value)  # Column B - QTY
            uom = uploaded_ws[f'C{row}'].value  # Column C - Unit of Measure
            unit_price = uploaded_ws[f'D{row}'].value  # Column D - Unit Price
            description = uploaded_ws[f'E{row}'].value  # Column E - Description
            upc = uploaded_ws[f'F{row}'].value  # Column F - UPC
            vendor_part = uploaded_ws[f'G{row}'].value  # Column G - Vendor Part
            sku = uploaded_ws[f'I{row}'].value  # Column I - SKU
            
            # Skip if no item number (empty row)
            if not item_no:
                row += 1
                continue
                
            # Calculate number of cartons needed
            if upc in counts:
                items_per_case = counts[upc]
                cartons_needed = qty // items_per_case
                remainder = qty % items_per_case
                
                print(f"Product {item_no}: QTY={qty}, Items per case={items_per_case}, Cartons={cartons_needed}, Remainder={remainder}")
                
                # Create one line per full carton
                for carton_num in range(cartons_needed):
                    carton_lines.append({
                        'item_no': item_no,
                        'qty': items_per_case,  # Full carton quantity
                        'uom': uom,
                        'unit_price': unit_price,
                        'description': description,
                        'vendor_part': vendor_part,
                        'sku': sku,
                        'upc': upc
                    })
                
                # If there's a remainder, create one more line for the partial carton
                if remainder > 0:
                    carton_lines.append({
                        'item_no': item_no,
                        'qty': remainder,  # Partial carton quantity
                        'uom': uom,
                        'unit_price': unit_price,
                        'description': description,
                        'vendor_part': vendor_part,
                        'sku': sku,
                        'upc': upc
                    })
            else:
                print(f"Warning: UPC {upc} not found in counts dictionary for item {item_no}")
                # If UPC not found, create one line with original quantity
                carton_lines.append({
                    'item_no': item_no,
                    'qty': qty,
                    'uom': uom,
                    'unit_price': unit_price,
                    'description': description,
                    'vendor_part': vendor_part,
                    'sku': sku,
                    'upc': upc
                })
            
            row += 1
                
        except (ValueError, TypeError) as e:
            print(f"Error processing row {row}: {e}")
            row += 1
            continue
    
    # Write the carton-based lines to the destination worksheet
    for i, line_data in enumerate(carton_lines):
        target_row = target_start_row + i
        
        # Write data to the appropriate columns
        source_ws[f'A{target_row}'] = line_data['item_no']  # Item number
        source_ws[f'F{target_row}'] = line_data['upc']      # UPC
        source_ws[f'G{target_row}'] = line_data['sku']      # SKU
        source_ws[f'H{target_row}'] = line_data['vendor_part']  # Vendor Part
        source_ws[f'I{target_row}'] = line_data['qty']      # QTY (now per carton)
        source_ws[f'J{target_row}'] = line_data['uom']      # Unit of Measure
        source_ws[f'K{target_row}'] = line_data['description']  # Description
        
        print(f"Created carton line {i+1}: Item {line_data['item_no']}, QTY {line_data['qty']}, UPC {line_data['upc']}")
    
    return len(carton_lines)  # Return the number of carton lines created

def create_carton_based_lines(xls_sheet, source_ws, start_row=23, target_start_row=19):
    """
    Creates one line per carton instead of one line per product for XLS files.
    For each product, calculates how many cartons are needed and creates that many lines.
    """
    carton_lines = []  # Store all the data for carton-based lines
    
    # Process each product row from the uploaded file
    for row_idx in range(start_row - 1, xls_sheet.nrows):  # start_row is 1-indexed, so subtract 1
        try:
            # Get product data from the uploaded file
            item_no = xls_sheet.cell_value(row_idx, 0)  # Column A - Item number
            qty = int(xls_sheet.cell_value(row_idx, 1))  # Column B - QTY
            uom = xls_sheet.cell_value(row_idx, 2)  # Column C - Unit of Measure
            unit_price = xls_sheet.cell_value(row_idx, 3)  # Column D - Unit Price
            description = xls_sheet.cell_value(row_idx, 4)  # Column E - Description
            vendor_part = xls_sheet.cell_value(row_idx, 6)  # Column G - Vendor Part
            sku = xls_sheet.cell_value(row_idx, 8)  # Column I - SKU
            upc = xls_sheet.cell_value(row_idx, 5)  # Column F - UPC
            
            # Skip if no item number (empty row)
            if not item_no:
                continue
                
            # Calculate number of cartons needed
            if upc in counts:
                items_per_case = counts[upc]
                cartons_needed = qty // items_per_case
                remainder = qty % items_per_case
                
                print(f"Product {item_no}: QTY={qty}, Items per case={items_per_case}, Cartons={cartons_needed}, Remainder={remainder}")
                
                # Create one line per full carton
                for carton_num in range(cartons_needed):
                    carton_lines.append({
                        'item_no': item_no,
                        'qty': items_per_case,  # Full carton quantity
                        'uom': uom,
                        'unit_price': unit_price,
                        'description': description,
                        'vendor_part': vendor_part,
                        'sku': sku,
                        'upc': upc
                    })
                
                # If there's a remainder, create one more line for the partial carton
                if remainder > 0:
                    carton_lines.append({
                        'item_no': item_no,
                        'qty': remainder,  # Partial carton quantity
                        'uom': uom,
                        'unit_price': unit_price,
                        'description': description,
                        'vendor_part': vendor_part,
                        'sku': sku,
                        'upc': upc
                    })
            else:
                print(f"Warning: UPC {upc} not found in counts dictionary for item {item_no}")
                # If UPC not found, create one line with original quantity
                carton_lines.append({
                    'item_no': item_no,
                    'qty': qty,
                    'uom': uom,
                    'unit_price': unit_price,
                    'description': description,
                    'vendor_part': vendor_part,
                    'sku': sku,
                    'upc': upc
                })
                
        except (ValueError, IndexError) as e:
            print(f"Error processing row {row_idx + 1}: {e}")
            continue
    
    # Write the carton-based lines to the destination worksheet
    for i, line_data in enumerate(carton_lines):
        target_row = target_start_row + i
        
        # Write data to the appropriate columns
        source_ws[f'A{target_row}'] = line_data['item_no']  # Item number
        source_ws[f'F{target_row}'] = line_data['upc']      # UPC
        source_ws[f'G{target_row}'] = line_data['sku']      # SKU
        source_ws[f'H{target_row}'] = line_data['vendor_part']  # Vendor Part
        source_ws[f'I{target_row}'] = line_data['qty']      # QTY (now per carton)
        source_ws[f'J{target_row}'] = line_data['uom']      # Unit of Measure
        source_ws[f'K{target_row}'] = line_data['description']  # Description
        
        print(f"Created carton line {i+1}: Item {line_data['item_no']}, QTY {line_data['qty']}, UPC {line_data['upc']}")
    
    return len(carton_lines)  # Return the number of carton lines created

def oneToMany_xlsx(uploaded_ws, source_ws, row, col, target_column, start_row, column_length):
    """
    Copies a value from one specific cell in XLSX and pastes it into multiple rows in a target column.
    """
    try:
        # Convert 1-indexed to 0-indexed for openpyxl
        cell_value = uploaded_ws.cell(row=row, column=col+1).value
        print(f"Copying value '{cell_value}' from ({row}, {col + 1})")

        # Ensure exactly `column_length` rows are pasted without overshooting
        for i in range(column_length):
            current_row = start_row + i  # Adjust to paste into the correct row
            source_ws[f'{target_column}{current_row}'] = cell_value
            print(f"Pasting '{cell_value}' into {target_column}{current_row}")

    except Exception as e:
        print(f"Unexpected error: {str(e)}")

# Function to convert .xls to .xlsx and transfer data to a backup of ASN copy
def convert_xls_data(uploaded_file, dest_file):
    xls_book = xlrd.open_workbook(uploaded_file)
    source_wb = load_workbook(source_asn_xlsx)
    source_ws = source_wb.active
    xls_sheet = xls_book.sheet_by_index(0)

    # Mapping uploaded cells to copy cells
    data_map = {
        (17, 1): 'E4',   # 'B18' -> (18, 2) name
        (17, 3): 'E5',   # 'D18' -> (18, 4) number
        (17, 4): 'E6',   # 'E18' -> (18, 5) add 1
        (17, 9): 'E7',   # 'J18' -> (18, 10) city
        (17, 10): 'E8',  # 'K18' -> (18, 11) State
        (17, 11): 'E9',  # 'L18' -> (18, 12) Zip
        (9, 2): 'E14'    # 'C10' -> (10, 3) delivery date
    }

    for (row, col), copy_cell in data_map.items():
        value = xls_sheet.cell_value(row, col)
        source_ws[copy_cell] = value

    # Manually assign values for cells that don’t come from the xls_sheet
    source_ws['B12'] = "FEDG"
    source_ws['B13'] = "Fedex"
    
    # Set current date in E11 and highlight it
    current_date = datetime.datetime.now().strftime("%m/%d/%Y")
    source_ws['E11'] = current_date
    
    # Highlight cell E11 with yellow background
    from openpyxl.styles import PatternFill
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    source_ws['E11'].fill = yellow_fill

    # Debugging: Print the total number of rows and columns in the uploaded sheet
    total_rows = xls_sheet.nrows
    total_cols = xls_sheet.ncols
    print(f"Total rows: {total_rows}, Total columns: {total_cols}")

    # Ensure you don't go out of bounds
    if total_rows < 23:
        print(f"Error: Not enough rows to start processing from row 23.")
        return

    # Create carton-based lines instead of product-based lines
    carton_lines_count = create_carton_based_lines(xls_sheet, source_ws, start_row=23, target_start_row=19)
    
    print(f"Total carton lines created: {carton_lines_count}")
    
    # Additional copy and paste operations with helper functions
    oneToMany(xls_sheet, source_ws, row=3, col=2, target_column='B', start_row=19, column_length=carton_lines_count)  # PO
    oneToMany(xls_sheet, source_ws, row=3, col=7, target_column='C', start_row=19, column_length=carton_lines_count)  # PO Date

    typedValue(source_ws, static_value="NA", target_column='D', start_row=19, column_length=carton_lines_count)

    # Set total cartons in B15 and "C" in E13
    source_ws['B15'] = carton_lines_count
    source_ws['E13'] = "C"

    format_cells_as_text(source_ws)
    format_cells_as_text(source_ws)
    align_cells_left(source_ws)
    align_cells_left(source_ws)

    # Save the updated file
    try:
        source_wb.save(dest_file)
        print(f"File saved successfully as {dest_file}.")
    except Exception as e:
        print(f"Error saving file: {str(e)}")


def process_MurdochsASN(file_path):
    """Main function to process Murdochs ASN files."""
    current_date = datetime.datetime.now().strftime("%m.%d.%Y")
    # Determine if the file is XLSX or XLS and extract the PO number
    if file_path.endswith('.xlsx'):
        uploaded_wb = load_workbook(file_path)
        po_number = uploaded_wb.active['C4'].value
    elif file_path.endswith('.xls'):
        xls_book = xlrd.open_workbook(file_path)
        po_number = xls_book.sheet_by_index(0).cell_value(3, 2)

    # Define the backup file path in FINISHED_FOLDER with a Murdochs subfolder
    backup_file = os.path.join(FINISHED_FOLDER, f"Murdochs/Murdochs 856 ASN PO {po_number} {current_date}.xlsx")
    
    # Ensure the directory exists
    os.makedirs(os.path.dirname(backup_file), exist_ok=True)

    # Perform the copy or conversion based on file type
    if file_path.endswith('.xlsx'):
        copy_xlsx_data(file_path, backup_file)
    elif file_path.endswith('.xls'):
        convert_xls_data(file_path, backup_file)

    return backup_file, po_number
