from flask import Flask, request, jsonify, render_template
import os
from processors.Chewy import process_chewy
from processors.ChewyLabel import process_label

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def upload_form():
    return render_template('upload.html')  # Serve the form from an HTML file

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'message': 'No file part'}), 400

    file = request.files['file']
    company = request.form.get('company')  # Get selected company from dropdown

    if file.filename == '':
        return jsonify({'message': 'No selected file'}), 400

    if file:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        try:
            # Route file processing based on selected company
            if company == 'chewy':
                process_chewy(file_path)
                process_label(file_path)
            elif company == 'companyB':
                process_companyB(file_path)
            elif company == 'companyC':
                process_companyC(file_path)

        except Exception as e:
            print(f"Error: {e}")
            return jsonify({'message': f'Error processing file: {str(e)}'}), 500

        return jsonify({'message': 'File processed successfully!'})

if __name__ == '__main__':
    app.run(debug=True)
