from flask import Flask, request, jsonify, render_template
from openpyxl import load_workbook, Workbook
import os
import xlrd

app = Flask(__name__)

# Create upload folder
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Define the destination filenames for the copies
chewy_asn_copy = "Finished/Chewy/Chewy 856 ASN - Copy - Backup.xlsx"
chewy_label_copy = "Finished/Chewy/Chewy UCC128 Label Request - Copy - Backup.xlsx"

# Function to copy .xlsx files using openpyxl
def copy_xlsx(source_file, dest_file):
    workbook = load_workbook(source_file)
    workbook.save(dest_file)

# Function to convert .xls to .xlsx using xlrd and openpyxl
def convert_xls_to_xlsx(source_file, dest_file):
    xls_book = xlrd.open_workbook(source_file)
    xlsx_book = Workbook()
    sheet = xlsx_book.active
    xls_sheet = xls_book.sheet_by_index(0)
    
    for row_idx in range(xls_sheet.nrows):
        row = xls_sheet.row_values(row_idx)
        sheet.append(row)
    
    xlsx_book.save(dest_file)

# Serve the upload form on the root path
@app.route('/')
def upload_form():
    return '''
    <h2>Upload an Excel File</h2>
    <form id="uploadForm" enctype="multipart/form-data" method="POST" action="/upload">
        <input type="file" id="excelFile" name="file" accept=".xlsx, .xls" required>
        <button type="submit">Upload and Copy</button>
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

        if file.filename.endswith('.xlsx'):
            copy_xlsx(file_path, chewy_asn_copy)
        elif file.filename.endswith('.xls'):
            convert_xls_to_xlsx(file_path, chewy_label_copy)

        return jsonify({'message': 'File processed successfully!'})

if __name__ == '__main__':
    app.run(debug=True)
