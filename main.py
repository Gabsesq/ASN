from flask import Flask, request, render_template, send_from_directory
import os
from processors.Chewy import process_chewy
from processors.ChewyLabel import process_label
from processors.TSC import process_TSC
from processors.PetSupermarketASN import process_PetSuperASN
from processors.PetSupermarketLabel import process_PetSuperLabel
from processors.ThriveASN import process_ThriveASN
from processors.MurdochsASN import process_MurdochsASN

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
FINISHED_FOLDER = 'Finished/Chewy'  # Assuming processed files go into a subfolder of Finished

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
    company = request.form.get('company')  # Get selected company from dropdown

    if file.filename == '':
        return render_template('back.html', message='No selected file'), 400

    if file:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        try:
            processed_files = []  # List to store the paths of processed files

            # Route file processing based on selected company
            if company == 'chewy':
                # Call your processing functions
                process_chewy(file_path)
                process_label(file_path)
                
                # Add processed files to the list (assuming the filenames are based on PO numbers)
                for filename in os.listdir(FINISHED_FOLDER):
                    if filename.endswith('.xlsx'):
                        processed_files.append(filename)

            elif company == 'TSC':
                process_TSC(file_path)
                processed_files.append(file.filename)  # Append the processed file to list
            elif company == 'PetSupermarket':
                process_PetSuperLabel(file_path)
                process_PetSuperASN(file_path)
                processed_files.append(file.filename)  # Append the processed file to list
            elif company == 'Thrive':
                process_ThriveASN(file_path)
                processed_files.append(file.filename)  # Append the processed file to list
            elif company == 'Murdochs':
                process_MurdochsASN(file_path)
                processed_files.append(file.filename)  # Append the processed file to list



        except Exception as e:
            print(f"Error: {e}")
            return render_template('back.html', message=f'Error processing file: {str(e)}'), 500

        # Return the success page with the list of processed files
        return render_template('back.html', message='File processed successfully!', files=processed_files)

# Serve files from the Finished folder
@app.route('/finished/<filename>')
def finished_file(filename):
    return send_from_directory(FINISHED_FOLDER, filename)

if __name__ == '__main__':
    app.run(debug=True)
