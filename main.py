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
from processors.Chewy20 import process_Chewy20

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
    company = request.form.get('company')

    # Check for file uploads
    asn_file = request.files.get('asn_file_1')
    label_file = request.files.get('label_file')

    if not asn_file or (company == "Chewy20" and not label_file):
        return render_template('back.html', message="Missing required files for processing."), 400

    # Save files locally
    asn_file_path = os.path.join(UPLOAD_FOLDER, asn_file.filename)
    asn_file.save(asn_file_path)

    label_file_path = None
    if label_file:
        label_file_path = os.path.join(UPLOAD_FOLDER, label_file.filename)
        label_file.save(label_file_path)

    try:
        # Process based on company selection
        if company == "Chewy20":
            processed_file, po_number = process_Chewy20(asn_file_path, label_file_path)
            processed_files = [processed_file]
        elif company == "TSC":
            processed_file, po_number = process_TSC(asn_file_path)
            processed_files = [processed_file]
        else:
            asn_processor = globals().get(f"process_{company}ASN")
            label_processor = globals().get(f"process_{company}Label")
            if not asn_processor or not label_processor:
                raise ValueError("Invalid company selected.")
            asn_output, po_number = asn_processor(asn_file_path)
            label_output, _ = label_processor(asn_file_path)
            processed_files = [asn_output, label_output]

        # Render success page with download links
        return render_template(
            'back.html',
            message="Files processed successfully!",
            processed_files=processed_files,
            company=company,
            po_number=po_number,
        )
    except Exception as e:
        print(f"Error processing files: {e}")
        return render_template('back.html', message=f"Error processing files: {str(e)}"), 500



@app.route('/download/<path:file_path>')
def download_file(file_path):
    """Serve the processed file to the user."""
    full_file_path = resource_path(file_path)

    if not os.path.exists(full_file_path):
        print(f"File not found: {full_file_path}")
        return f"File not found", 404

    return send_file(full_file_path, as_attachment=True)

@app.route('/shutdown', methods=['POST'])
def shutdown():
    """Shut down the server."""
    shutdown_function = request.environ.get('werkzeug.server.shutdown')
    if shutdown_function:
        shutdown_function()
    return 'Server shutting down...'

if __name__ == '__main__':
    # Automatically open the default web browser to the Flask app’s URL
    port = 5000  # Set your preferred port if needed
    url = f"http://127.0.0.1:{port}"
    webbrowser.open(url)  # Open the browser to this URL
    
    app.run(host="0.0.0.0", port=port)  # Start the Flask app