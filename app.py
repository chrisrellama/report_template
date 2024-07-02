from flask import Flask, request, render_template, send_from_directory
import pandas as pd
from datetime import date
import os
from docxtpl import DocxTemplate

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    excel_file = request.files['excel_file']
    template_file = request.files['template_file']
    
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_file.filename)
    
    excel_file.save(excel_path)
    template_file.save(template_path)
    
    report_files = generate_reports(excel_path, template_path)
    
    return render_template('upload.html', reports=report_files, basename=os.path.basename)

def generate_reports(excel_path, template_path):
    excel_data = pd.read_excel(excel_path)
    today = date.today().strftime("%Y-%m-%d")
    report_files = []

    for index, row in excel_data.iterrows():
        doc = DocxTemplate(template_path)
        context = {column: str(row[column]) for column in excel_data.columns}
        context['report_date'] = today
        
        output_file = os.path.join(app.config['UPLOAD_FOLDER'], f'report_{index + 1}_{today}.docx')
        doc.render(context)
        doc.save(output_file)
        report_files.append(output_file)
    
    return report_files

@app.route('/uploads/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
