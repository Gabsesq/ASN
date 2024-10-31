from flask import Flask, request, render_template, send_file
import os
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

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
FINISHED_FOLDER = 'Finished'

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

            # Process files based on the company
            if company == 'TSC':
                # Only process one file for TSC
                processed_file_path = process_TSC(file_path)
                processed_files = [processed_file_path]  # Single file for TSC
            else:
                # Process two files for all other companies
                asn_file_path = globals()[f"process_{company}ASN"](file_path)
                label_file_path = globals()[f"process_{company}Label"](file_path)
                processed_files = [asn_file_path, label_file_path]

            # Render success page with download links for processed files
            return render_template(
                'back.html',
                message='File processed successfully!',
                processed_files=processed_files,
                company=company
            )

        except Exception as e:
            print(f"Error: {e}")
            return render_template('back.html', message=f'Error processing file: {str(e)}'), 500

@app.route('/download/<path:file_path>')
def download_file(file_path):
    """Serve the processed file to the user."""
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return f"File not found", 404

    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
