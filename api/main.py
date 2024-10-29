from flask import Flask, request, render_template, send_from_directory
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

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
FINISHED_FOLDER = 'Finished'  # Assuming processed files go into a subfolder of Finished

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
            # Determine the output folder based on the company
            output_folder = os.path.join(FINISHED_FOLDER, company)
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            # Process the file based on the selected company
            if company == 'chewy':
                process_chewy(file_path)
                process_label(file_path)
            elif company == 'TSC':
                process_TSC(file_path)
            elif company == 'PetSupermarket':
                process_PetSuperLabel(file_path)
                process_PetSuperASN(file_path)
            elif company == 'Thrive':
                process_ThriveASN(file_path)
                process_ThriveLabel(file_path)
            elif company == 'Murdochs':
                process_MurdochsASN(file_path)
                process_MurdochsLabel(file_path)
            elif company == 'Scheels':
                process_ScheelsASN(file_path)

            # Collect all processed files from the output folder
            processed_files = [
                filename for filename in os.listdir(output_folder) if filename.endswith('.xlsx')
            ]

            if not processed_files:
                return render_template('back.html', message='No processed files found.', files=[])

            # Render the success page with the list of processed files and company
            return render_template(
                'back.html',
                message='File processed successfully!',
                files=processed_files,
                company=company
            )

        except Exception as e:
            print(f"Error: {e}")
            return render_template('back.html', message=f'Error processing file: {str(e)}'), 500


@app.route('/finished/<company>/<filename>')
def finished_file(company, filename):
    """Serve the processed file to the user."""
    try:
        output_folder = os.path.join(FINISHED_FOLDER, company)
        file_path = os.path.join(output_folder, filename)

        # Debugging: Check if the file exists
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return f"File {filename} not found", 404

        print(f"Serving file: {file_path}")
        return send_from_directory(output_folder, filename, as_attachment=True)

    except Exception as e:
        print(f"Error serving file: {str(e)}")
        return f"Error serving file: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)
