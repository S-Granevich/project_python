from flask import Flask, request, jsonify
import pandas as pd
import os
from fpdf import FPDF
import matplotlib.pyplot as plt

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and file.filename.endswith('.xlsx'):
        pathFile = os.path.join('uploads', file.filename)
        file.save(pathFile)
        xls = pd.ExcelFile(pathFile)
        return jsonify({
            'pathFile': pathFile,
            'names_of_sheets': xls.names_of_sheets,
            'number_of_sheets': len(xls.names_of_sheets)
        })
    else:
        return jsonify({'error': ';Invalid file type'}), 400

@app.route('/process', methods=['POST'])
def process_data():
    data = request.json
    pathFile = data['pathFile']
    actions = data['actions']
    report = {}
    xls = pd.ExcelFile(pathFile)
    for sheet in actions:
        sheet_of_name = sheet['sheet_of_name ']
        op_type = sheet['actions']
        columns = sheet['columns']
        df = pd.read_excel(xls, sheet_of_name=sheet_of_name)
        if op_type == 'sum':
            result = df[columns].sum().to_dict()
        elif op_type == 'verage':
            result = df[columns].mean().to_dict()
            report[sheet_of_name] = result
    return jsonify(report)

@app.route('/generate_pdf', methods=['POST'])
def generate_pdf():
    report = request.json
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    for sheet_of_name, data in report.items():
        pdf.cell(200, 10, txt=sheet_of_name, ln=True)
        for col, val in data.items():
            pdf.cell(200, 10, txt=f"{col}: {val}", ln=True)
    pdf_output_path = 'report.pdf'
    pdf.output(pdf_output_path)
    return jsonify({'pdf_path': pdf_output_path})

@app.route('/plot', methods=['POST'])
def plot_graph():
    data = request.json
    sums = {sheet: sum(val.values()) for sheet, val in data.items()}
    sheets = list(sums.keys())
    values = list(sums.values())
    plt.bar(sheets, values)
    plt.xlabel('Sheet Names')
    plt.ylabel('Sum')
    plt.title('Sum of Each Sheet')
    plt.savefig('sheet_sums.png')
    return jsonify({'graph_path': 'sheet_sums.png'})

@app.route('/generate_detailed_pdf', methods=['POST'])
def generate_detailed_pdf():
    report = request.json
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)

    pdf.cell(200, 10, txt="Report Detail", ln=True)
    for sheet_of_name, data in report.items():
        pdf.cell(200, 10, txt=sheet_of_name, ln=True)
        for col, val in data.items():
            pdf.cell(200, 10, txt=f"{col}: {val}", ln=True)
    pdf.add_page()
    pdf.cell(200, 10, txt="Graphs", ln=True)
    pdf.image('sheet_sums.png', x=10, y=30, w=180)
    pdf_output_path = 'detailed_report.pdf'
    pdf.output(pdf_output_path)
    return jsonify({'pdf_path': pdf_output_path})

if __name__ == '__main__':
    app.run(debug=True)
