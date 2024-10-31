from flask import Flask, request, render_template, send_file
import os
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

print("Current Working Directory:", os.getcwd())

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
            # Process the file based on the selected company
            if company == 'chewy':
                processed_file_path = process_chewy(file_path)
                process_label(file_path)  # Ensure both files get processed if needed
            elif company == 'TSC':
                processed_file_path = process_TSC(file_path)
            elif company == 'PetSupermarket':
                processed_file_path = process_PetSuperASN(file_path)
                process_PetSuperLabel(file_path)
            elif company == 'Thrive':
                processed_file_path = process_ThriveASN(file_path)
                process_ThriveLabel(file_path)
            elif company == 'Murdochs':
                processed_file_path = process_MurdochsASN(file_path)
                process_MurdochsLabel(file_path)
            elif company == 'Scheels':
                processed_file_path = process_ScheelsASN(file_path)
            else:
                return render_template('back.html', message='Invalid company selected'), 400

            # Render success page with a download link for the processed file
            processed_filename = os.path.basename(processed_file_path)
            return render_template(
                'back.html',
                message='File processed successfully!',
                processed_filename=processed_filename,
                company=company
            )

        except Exception as e:
            print(f"Error: {e}")
            return render_template('back.html', message=f'Error processing file: {str(e)}'), 500

@app.route('/download/<company>/<processed_filename>')
def download_file(company, processed_filename):
    """Serve the processed file to the user."""
    # Construct the path for the processed file in the Finished folder
    file_path = os.path.join(FINISHED_FOLDER, company, processed_filename)

    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return f"File {processed_filename} not found", 404

    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
