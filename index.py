from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
from docx import Document
from fpdf import FPDF
from PIL import Image
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Conversão de Word para PDF
def convert_word_to_pdf(word_path, pdf_path):
    doc = Document(word_path)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    for paragraph in doc.paragraphs:
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, paragraph.text)

    pdf.output(pdf_path)

# Conversão de PNG para PDF
def convert_png_to_pdf(png_path, pdf_path):
    image = Image.open(png_path)
    rgb_image = image.convert('RGB')
    rgb_image.save(pdf_path)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']
    if file.filename == '':
        return "No selected file"

    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    name, ext = os.path.splitext(filename)
    ext = ext.lower()
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{name}.pdf")

    if ext == '.docx':
        convert_word_to_pdf(file_path, output_path)
    elif ext == '.png':
        convert_png_to_pdf(file_path, output_path)
    else:
        return "Unsupported file format"

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
