from flask import Flask, request, render_template, send_file
import os
import sys
import webbrowser
from openpyxl import load_workbook
import xlrd
from datetime import date, timedelta, datetime
from processors.ChewyASN import process_ChewyASN
from processors.ChewyLabel import process_ChewyLabel
from processors.TSC import process_TSC
from processors.TSCISASN import process_TSCISASN
from processors.TSCISLabel import process_TSCISLabel
from processors.PetSupermarketASN import process_PetSupermarketASN
from processors.PetSupermarketLabel import process_PetSupermarketLabel
from processors.ThriveASN import process_ThriveASN
from processors.ThriveLabel import process_ThriveLabel
from processors.MurdochsASN import process_MurdochsASN
from processors.MurdochsLabel import process_MurdochsLabel
from processors.ScheelsASN import process_ScheelsASN
from processors.ScheelsLabel import process_ScheelsLabel
from ExcelHelpers import resource_path, UPLOAD_FOLDER, FINISHED_FOLDER
from calendar_helpers import get_ship_date_recommendation, format_ship_date_for_calendar, get_calendar_event_details

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

# Create TSCIS folder in Finished directory
tscis_folder = os.path.join(FINISHED_FOLDER, 'TSCIS')
if not os.path.exists(tscis_folder):
    os.makedirs(tscis_folder)

def get_company_from_excel(file_path):
    """Reads cell A2 from an Excel file to determine the company."""
    try:
        company_string = None
        if file_path.endswith('.xlsx'):
            workbook = load_workbook(file_path)
            sheet = workbook.active
            company_string = sheet['A2'].value
        elif file_path.endswith('.xls'):
            workbook = xlrd.open_workbook(file_path)
            sheet = workbook.sheet_by_index(0)
            company_string = sheet.cell_value(1, 0)  # Row 2 (index 1), Column A (index 0)
        else:
            return None # Unsupported file type

        if not company_string:
            return None

        # Determine company based on keywords
        if "Chewy" in company_string:
            return "Chewy"
        if "Tractor Supply IS" in company_string:
            return "TSCIS"
        elif "Tractor Supply" in company_string:
            return "TSC"
        if "Pet Supermarket" in company_string:
            return "Pet Supermarket"
        if "Thrive" in company_string:
            return "Thrive"
        if "Murdoch" in company_string:
            return "Murdochs"
        if "Scheels" in company_string:
            return "Scheels"
        if "TSC" in company_string:
            return "TSC"
        
        return None  # No matching company found
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

@app.route('/')
def upload_form():
    return render_template('upload.html')  # Serve the form from an HTML file

@app.route('/calendar')
def calendar_view():
    return render_template('calendar.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Check for file uploads
        asn_file = request.files.get('asn_file_1')
        if not asn_file:
            print("No ASN file uploaded")
            return render_template('back.html', message="Missing ASN file."), 400

        # Save files locally
        asn_file_path = os.path.join(UPLOAD_FOLDER, asn_file.filename)
        asn_file.save(asn_file_path)
        print(f"Saved ASN file to: {asn_file_path}")

        # Determine company from Excel file
        company = get_company_from_excel(asn_file_path)
        if not company:
            return render_template('back.html', message="Could not determine company from Excel file."), 400
        
        # Add debug print
        print(f"Processing company: {company}")

        try:
            if company == "Chewy":
                asn_output, po_number = process_ChewyASN(asn_file_path)
                label_output, _ = process_ChewyLabel(asn_file_path)
                processed_files = [asn_output, label_output]
            elif company == "TSC":
                output, po_number = process_TSC(asn_file_path)
                processed_files = [output]
            elif company == "TSCIS":
                print("Processing TSCIS files...")
                # Process both ASN and Label files
                asn_output, po_number = process_TSCISASN(asn_file_path)
                print(f"ASN processed: {asn_output}")
                label_output, _ = process_TSCISLabel(asn_file_path)
                print(f"Label processed: {label_output}")
                processed_files = [asn_output, label_output]
            elif company == "Pet Supermarket":
                asn_output, po_number = process_PetSupermarketASN(asn_file_path)
                label_output, _ = process_PetSupermarketLabel(asn_file_path)
                processed_files = [asn_output, label_output]
            elif company == "Thrive":
                asn_output, po_number = process_ThriveASN(asn_file_path)
                label_output, _ = process_ThriveLabel(asn_file_path)
                processed_files = [asn_output, label_output]
            elif company == "Murdochs":
                asn_output, po_number = process_MurdochsASN(asn_file_path)
                label_output, _ = process_MurdochsLabel(asn_file_path)
                processed_files = [asn_output, label_output]
            elif company == "Scheels":
                asn_output, po_number = process_ScheelsASN(asn_file_path)
                label_output, _ = process_ScheelsLabel(asn_file_path)
                processed_files = [asn_output, label_output]
            else:
                return render_template('back.html', message=f"Unknown company: {company}"), 400

            # Get enhanced calendar event details
            calendar_event = get_calendar_event_details(company, asn_file_path, po_number)
            ship_date_str = calendar_event.get('ship_date')

            return render_template(
                'back.html',
                message="Files processed successfully!",
                processed_files=processed_files,
                company=company,
                po_number=po_number,
                ship_date=ship_date_str,
                calendar_event=calendar_event
            )

        except Exception as e:
            print(f"Error processing files: {str(e)}")
            return render_template('back.html', message=f"Error processing files: {str(e)}"), 500

    except Exception as e:
        print(f"Upload error: {str(e)}")
        return render_template('back.html', message=f"Upload error: {str(e)}"), 500

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
    # Automatically open the default web browser to the Flask app's URL
    port = 5000  # Set your preferred port if needed
    url = f"http://127.0.0.1:{port}"
    webbrowser.open(url)  # Open the browser to this URL
    
    app.run(host="0.0.0.0", port=port)  # Start the Flask app