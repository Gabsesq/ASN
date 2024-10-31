from flask import Flask, request, render_template, send_file, send_from_directory
import os
import zipfile
from processors.Chewy import process_chewy
from processors.ChewyLabel import process_label
from processors.TSC import process_TSC
from processors.PetSupermarketASN import process_PetSuperASN
from processors.PetSupermarketLabel import process_PetSuperLabel
from processors.ThriveASN import process_ThriveASN
from processors.ThriveLabel import process_ThriveLabel
from processors.MurdochsASN import process_MurdochsASN
from processors.MurdochsLabel import process_MurdochsLabel
from processors.ScheelsASN import process_ScheelsASN

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
FINISHED_FOLDER = 'Finished'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(FINISHED_FOLDER):
    os.makedirs(FINISHED_FOLDER)

def zip_folder(folder_path):
    """Create a ZIP file of the given folder."""
    zip_filename = os.path.basename(folder_path) + ".zip"  # Name the ZIP file after the PO number
    zip_path = os.path.join(folder_path, zip_filename)  # Save ZIP inside the PO folder

    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                zipf.write(os.path.join(root, file), 
                           os.path.relpath(os.path.join(root, file), folder_path))
    return zip_path  # Return full path to the ZIP file

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
            # Process the file based on the selected company
            if company == 'chewy':
                folder_path = process_chewy(file_path)  # Function should return the folder path
                process_label(file_path)
            elif company == 'TSC':
                folder_path = process_TSC(file_path)
            elif company == 'PetSupermarket':
                folder_path = process_PetSuperASN(file_path)
                process_PetSuperLabel(file_path)
            elif company == 'Thrive':
                folder_path = process_ThriveASN(file_path)
                process_ThriveLabel(file_path)
            elif company == 'Murdochs':
                folder_path = process_MurdochsASN(file_path)
                process_MurdochsLabel(file_path)
            elif company == 'Scheels':
                folder_path = process_ScheelsASN(file_path)

            # Zip the PO folder and get its path
            zip_path = zip_folder(folder_path)
            zip_filename = os.path.basename(zip_path)  # Extract the ZIP filename

            # Render success page with download link to the ZIP file
            return render_template(
                'back.html',
                message='Files processed successfully!',
                zip_filename=zip_filename,
                company=company,
                po_number=os.path.basename(folder_path),  # Pass PO number
                files=os.listdir(folder_path)  # List files in the folder
            )

        except Exception as e:
            print(f"Error: {e}")
            return render_template('back.html', message=f'Error processing file: {str(e)}'), 500

@app.route('/download/<company>/<po_number>/<zip_filename>')
def download_folder(company, po_number, zip_filename):
    """Serve the zipped folder to the user."""
    # Correct path: Start from the `Finished` folder
    zip_path = os.path.join(FINISHED_FOLDER, company, po_number, zip_filename)
    
    if not os.path.exists(zip_path):
        print(f"File not found: {zip_path}")
        return f"File {zip_filename} not found", 404

    return send_file(zip_path, as_attachment=True)



if __name__ == '__main__':
    app.run(debug=True)
