from flask import Flask, request, render_template, send_file
import os
import sys
import webbrowser
from processors.ChewyASN import process_ChewyASN
from processors.ChewyLabel import process_ChewyLabel
from processors.TSC import process_TSC
from processors.PetSupermarketASN import process_PetSupermarketASN
from processors.PetSupermarketLabel import process_PetSupermarketLabel
from processors.ThriveASN import process_ThriveASN
from processors.ThriveLabel import process_ThriveLabel
from processors.MurdochsASN import process_MurdochsASN
from processors.MurdochsLabel import process_MurdochsLabel
from processors.ScheelsASN import process_ScheelsASN
from processors.ScheelsLabel import process_ScheelsLabel
from ExcelHelpers import resource_path, UPLOAD_FOLDER, FINISHED_FOLDER

if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    print('running in a PyInstaller bundle')
else:
    print('running in a normal Python process')

# Initialize Flask with an absolute path to the templates folder
app = Flask(__name__, template_folder=resource_path("templates"))

UPLOAD_FOLDER = resource_path('uploads')
FINISHED_FOLDER = resource_path('Finished')

# Ensure necessary folders exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(FINISHED_FOLDER):
    os.makedirs(FINISHED_FOLDER)

@app.route('/')
def upload_form():
    return render_template('upload.html')  # Serve the form from an HTML file

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return render_template('back.html', message='No file part'), 400

    file = request.files['file']
    company = request.form.get('company')

    if file.filename == '':
        return render_template('back.html', message='No selected file'), 400

    if file:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        try:
            processed_files = []  # Initialize an empty list for file paths
            po_number = None       # Initialize PO number

            # Process files based on the company
            if company == 'TSC':
                # Only process one file for TSC
                processed_file_path, po_number = process_TSC(file_path)
                processed_files = [processed_file_path]  # Single file for TSC
            else:
                # Process two files for all other companies
                asn_file_path, po_number = globals()[f"process_{company}ASN"](file_path)
                label_file_path, _ = globals()[f"process_{company}Label"](file_path)
                processed_files = [asn_file_path, label_file_path]

            # Adjust paths for processed files to use resource_path
            processed_files = [resource_path(file[0]) if isinstance(file, tuple) else resource_path(file) for file in processed_files]

            # Render success page with download links for processed files
            return render_template(
                'back.html',
                message='File processed successfully!',
                processed_files=processed_files,
                company=company,
                po_number=po_number  # Pass the PO number to the template
            )

        except Exception as e:
            print(f"Error: {e}")
            return render_template('back.html', message=f'Error processing file: {str(e)}'), 500

@app.route('/download/<path:file_path>')
def download_file(file_path):
    """Serve the processed file to the user."""
    full_file_path = resource_path(file_path)

    if not os.path.exists(full_file_path):
        print(f"File not found: {full_file_path}")
        return f"File not found", 404

    return send_file(full_file_path, as_attachment=True)

if __name__ == '__main__':
    # Automatically open the default web browser to the Flask app’s URL
    port = 5000  # Set your preferred port if needed
    url = f"http://127.0.0.1:{port}"
    webbrowser.open(url)  # Open the browser to this URL
    
    app.run(host="0.0.0.0", port=port)  # Start the Flask app