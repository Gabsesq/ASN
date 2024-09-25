from flask import Flask, request, jsonify, render_template
from openpyxl import load_workbook, Workbook
import os
import xlrd

app = Flask(__name__)

# Create upload folder
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Define source files and destination copies
source_asn_xlsx = "assets/Chewy/Chewy 856 ASN - Copy.xlsx"
source_label_xls = "assets/Chewy/Chewy UCC128 Label Request - Copy.xls"
asn_copy_backup = "Finished/Chewy/Chewy 856 ASN - Copy - Backup.xlsx"
label_copy_backup = "Finished/Chewy/Chewy UCC128 Label Request - Copy - Backup.xlsx"

# Function to copy data from uploaded .xlsx file to specific cells in the ASN .xlsx backup
def copy_xlsx_data(uploaded_file, dest_file):
    uploaded_wb = load_workbook(uploaded_file)
    source_wb = load_workbook(source_asn_xlsx)
    uploaded_ws = uploaded_wb.active
    source_ws = source_wb.active

    # Mapping uploaded cells to copy cells
    data_map = {
        'B16': 'E3',   # uploaded B16 -> copy E3
        'D16': 'E4',   # uploaded D16 -> copy E4
        'E16': 'E5',   # uploaded E16 -> copy E5
        'J16': 'E6',   # uploaded J16 -> copy E6
        'K16': 'E7',   # uploaded K16 -> copy E7
        'L16': 'E8',   # uploaded L16 -> copy E8
        'C8':  'E14',  # uploaded C8  -> copy E14
    }

    # Loop through the map and transfer data from the uploaded file to the backup copy
    for upload_cell, copy_cell in data_map.items():
        source_ws[copy_cell] = uploaded_ws[upload_cell].value

    # Save the updated copy
    source_wb.save(dest_file)

# Function to convert .xls to .xlsx and transfer data to a backup of ASN copy
def convert_xls_data(uploaded_file, dest_file):
    # Open .xls file using xlrd
    xls_book = xlrd.open_workbook(uploaded_file)
    source_wb = load_workbook(source_asn_xlsx)
    source_ws = source_wb.active
    xls_sheet = xls_book.sheet_by_index(0)

    # Mapping uploaded cells to copy cells
    data_map = {
        (15, 1): 'E3',   # uploaded B16 -> copy E3 (zero-indexed for xlrd)
        (15, 3): 'E4',   # uploaded D16 -> copy E4
        (15, 4): 'E5',   # uploaded E16 -> copy E5
        (15, 9): 'E6',   # uploaded J16 -> copy E6
        (15, 10): 'E7',  # uploaded K16 -> copy E7
        (15, 11): 'E8',  # uploaded L16 -> copy E8
        (7, 2): 'E14',   # uploaded C8 -> copy E14
    }

    # Transfer data from .xls to the backup copy
    for (row, col), copy_cell in data_map.items():
        value = xls_sheet.cell_value(row, col)
        source_ws[copy_cell] = value

    # Save the updated copy
    source_wb.save(dest_file)

# Serve the upload form on the root path
@app.route('/')
def upload_form():
    return '''
    <h2>Upload an Excel File</h2>
    <form id="uploadForm" enctype="multipart/form-data" method="POST" action="/upload">
        <input type="file" id="excelFile" name="file" accept=".xlsx, .xls" required>
        <button type="submit">Upload and Transfer Data</button>
    </form>
    '''

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'message': 'No file part'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'message': 'No selected file'}), 400

    if file:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        try:
            if file.filename.endswith('.xlsx'):
                # Copy the data to the backup of the ASN .xlsx file
                copy_xlsx_data(file_path, asn_copy_backup)
            elif file.filename.endswith('.xls'):
                # Convert the .xls and transfer data to the backup of the label .xls file
                convert_xls_data(file_path, asn_copy_backup)
        except Exception as e:
            return jsonify({'message': f'Error processing file: {str(e)}'}), 500

        return jsonify({'message': 'File processed and data transferred successfully!'})

if __name__ == '__main__':
    app.run(debug=True)
