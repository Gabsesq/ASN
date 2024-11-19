from openpyxl import load_workbook
import xlrd
import os

def process_Chewy20(asn_file_path, carton_label_file_path):
    # Load the ASN file
    asn_wb = load_workbook(asn_file_path)
    asn_ws = asn_wb.active

    # Parse carton label file
    carton_data = []
    if carton_label_file_path.endswith('.xls'):
        xls_book = xlrd.open_workbook(carton_label_file_path)
        xls_sheet = xls_book.sheet_by_index(0)

        # Extract carton label data starting from row 20
        for row_idx in range(19, xls_sheet.nrows):  # Start from row 20 (index 19)
            row = xls_sheet.row_values(row_idx)
            carton_label = str(row[3]).strip()  # Column D = index 3
            upc = str(row[4]).strip() if row[4] else None  # Column E = index 4
            vendor_part = str(row[5]).strip() if row[5] else None  # Column F = index 5
            sku = str(row[6]).strip() if row[6] else None  # Column G = index 6

            if carton_label:  # Only append valid rows
                carton_data.append({
                    "carton_label": carton_label,
                    "upc": upc,
                    "vendor_part": vendor_part,
                    "sku": sku,
                })
    elif carton_label_file_path.endswith('.xlsx'):
        carton_label_wb = load_workbook(carton_label_file_path)
        carton_label_ws = carton_label_wb.active

        # Extract carton label data starting from row 20
        for row in carton_label_ws.iter_rows(min_row=20, values_only=True):  # Start from row 20
            carton_label = str(row[3]).strip() if row[3] else None  # Column D
            upc = str(row[4]).strip() if row[4] else None  # Column E
            vendor_part = str(row[5]).strip() if row[5] else None  # Column F
            sku = str(row[6]).strip() if row[6] else None  # Column G

            if carton_label:  # Only append valid rows
                carton_data.append({
                    "carton_label": carton_label,
                    "upc": upc,
                    "vendor_part": vendor_part,
                    "sku": sku,
                })
    else:
        raise ValueError("Unsupported file format. Please upload a .xls or .xlsx file for the carton labels.")

    print("Parsed Carton Data:", carton_data)  # Debug parsed data

    # Match and update ASN file
    used_labels = set()  # Track used carton labels
    for asn_row in asn_ws.iter_rows(min_row=20, values_only=False):  # Start from row 20
        upc_cell, vendor_part_cell, sku_cell, pallet_label_cell = asn_row[4], asn_row[5], asn_row[6], asn_row[3]
        upc = str(upc_cell.value).strip() if upc_cell.value else None
        vendor_part = str(vendor_part_cell.value).strip() if vendor_part_cell.value else None
        sku = str(sku_cell.value).strip() if sku_cell.value else None

        if not (upc and vendor_part and sku):
            continue

        for carton in carton_data:
            if (
                carton["upc"] == upc and
                carton["vendor_part"] == vendor_part and
                carton["sku"] == sku and
                carton["carton_label"] not in used_labels
            ):
                pallet_label_cell.value = carton["carton_label"]  # Update the ASN file
                used_labels.add(carton["carton_label"])  # Mark label as used
                print(f"Matched Carton Label: {carton['carton_label']} -> Row {asn_row[0].row}")
                break

    output_path = os.path.join(os.path.dirname(asn_file_path), "Updated_Chewy_ASN.xlsx")
    asn_wb.save(output_path)
    print(f"Updated ASN file saved to: {output_path}")
    return output_path, None
