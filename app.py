# app.py

from flask import Flask, render_template, request, send_file
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import pandas as pd
import os
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    # Check if the post request has the file part
    if 'pdf_file' not in request.files:
        return "No file part"

    pdf_file = request.files['pdf_file']

    # Check if the file is allowed
    if pdf_file and allowed_file(pdf_file.filename):
        # Save the PDF file to the server
        pdf_filename = secure_filename(pdf_file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
        pdf_file.save(pdf_path)

        # Convert PDF to images using pdf2image
        images = convert_from_path(pdf_path, 300)

        # OCR each image and extract text
        text_data = ""
        for image in images:
            text_data += pytesseract.image_to_string(image, lang='eng')

        # Convert text to DataFrame and save as XLSX
        df = pd.DataFrame([text_data.split('\n')])
        xls_file_path = 'converted_file.xlsx'
        df.to_excel(xls_file_path, index=False, header=False)

        # Provide the option to download the converted file
        return send_file(xls_file_path, as_attachment=True)
    else:
        return "Invalid file format"

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True)
