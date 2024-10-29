from flask import Flask, request, render_template, send_from_directory
import os
import sys
from werkzeug.utils import secure_filename
from config import TMP_PATH  # Ensure TMP_PATH is imported

# Use /tmp for Vercel, or 'tmp/' for local testing
TMP_PATH = '/tmp' if os.environ.get('VERCEL') else 'tmp'

# Ensure the tmp/ folder exists locally (for testing)
if not os.path.exists(TMP_PATH):
    os.makedirs(TMP_PATH)

# Add the parent directory to Python's path to import processors
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import processing functions
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

# Initialize Flask app
app = Flask(__name__, template_folder="../templates")

@app.route('/')
def upload_form():
    """Serve the upload form."""
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file uploads and processing."""
    if 'file' not in request.files:
        return render_template('back.html', message='No file part'), 400

    file = request.files['file']
    company = request.form.get('company')

    if file.filename == '':
        return render_template('back.html', message='No selected file'), 400

    # Save the uploaded file to /tmp or tmp/ (local)
    filename = secure_filename(file.filename)
    uploaded_path = os.path.join(TMP_PATH, filename)

    try:
        file.save(uploaded_path)  # Save the uploaded file
    except Exception as e:
        return render_template('back.html', message=f'Error saving file: {str(e)}'), 500

    try:

        # Save the uploaded file
        file.save(uploaded_path)
        print(f"Uploaded file saved to: {uploaded_path}")
        
        # Initialize a list to store processed file paths
        processed_files = []

        # Process based on the selected company
        if company == 'chewy':
            file1 = process_chewy(uploaded_path)
            file2 = process_label(uploaded_path)
            processed_files.extend([file1, file2])
        elif company == 'TSC':
            file1 = process_TSC(uploaded_path)
            if file1:
                processed_files.append(file1)
        elif company == 'PetSupermarket':
            file1 = process_PetSuperLabel(uploaded_path)
            file2 = process_PetSuperASN(uploaded_path)
            processed_files.extend([file1, file2])
        elif company == 'Thrive':
            file1 = process_ThriveASN(uploaded_path)
            file2 = process_ThriveLabel(uploaded_path)
            processed_files.extend([file1, file2])
        elif company == 'Murdochs':
            file1 = process_MurdochsASN(uploaded_path)
            file2 = process_MurdochsLabel(uploaded_path)
            processed_files.extend([file1, file2])
        elif company == 'Scheels':
            file1 = process_ScheelsASN(uploaded_path)
            if file1:
                processed_files.append(file1)

        # Filter out any None or non-existent files
        valid_files = [f for f in processed_files if f and os.path.exists(f)]

        if not valid_files:
            return render_template('back.html', message='No processed files found.', files=[])

        # Extract filenames to pass to the template
        file_names = [os.path.basename(f) for f in valid_files]

        # Render the success page with the list of processed files
        return render_template(
            'back.html',
            message='File processed successfully!',
            files=file_names,
            company=company
        )

    except Exception as e:
        print(f"Error processing file: {e}")
        return render_template('back.html', message=f'Error processing file: {str(e)}'), 500

@app.route('/finished/<company>/<filename>')
def finished_file(company, filename):
    """Serve the processed file to the user."""
    try:
        # Look for the file in the company's tmp folder
        file_path = os.path.join(TMP_PATH, filename)

        if not os.path.exists(file_path):
            return f"File {filename} not found", 404

        return send_from_directory(TMP_PATH, filename, as_attachment=True)

    except Exception as e:
        print(f"Error serving file: {str(e)}")
        return f"Error serving file: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)
