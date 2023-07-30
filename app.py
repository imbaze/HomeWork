import os
from flask import Flask, request, render_template, send_file
from PyPDF2 import PdfReader
from docx import Document

app = Flask(__name__)

input_folder = "input"
output_folder = "output"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'pdf_file' not in request.files:
        return "No file part"

    pdf_file = request.files['pdf_file']

    if pdf_file.filename == '':
        return "No selected file"

    if pdf_file:
        pdf_filename = pdf_file.filename
        pdf_path = os.path.join(input_folder, pdf_filename)
        pdf_file.save(pdf_path)

        docx_filename = os.path.splitext(pdf_filename)[0] + '.docx'
        docx_path = os.path.join(output_folder, docx_filename)
        convert_pdf_to_docx(pdf_path, docx_path)

        return send_file(docx_path, as_attachment=True)
    else:
        return "File not supported"

def convert_pdf_to_docx(input_pdf, output_docx):
    pdf = PdfReader(input_pdf)

    doc = Document()
    for page in pdf.pages:
        text = page.extract_text()
        doc.add_paragraph(text)

    doc.save(output_docx)
    return

if __name__ == '__main__':
    app.run()