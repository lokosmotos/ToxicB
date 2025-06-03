from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import io
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Allowed file extensions
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_sheets', methods=['POST'])
def get_sheets():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Please upload an Excel file (.xlsx, .xls)'}), 400

    try:
        # Read Excel file
        xl = pd.ExcelFile(file)
        sheets = xl.sheet_names
        return jsonify({'sheets': sheets, 'default_sheet': 'audio' if 'audio' in [s.lower() for s in sheets] else sheets[0]})
    except Exception as e:
        return jsonify({'error': f'Error reading file: {str(e)}'}), 500

@app.route('/get_columns', methods=['POST'])
def get_columns():
    if 'file' not in request.files or not request.form.get('sheet'):
        return jsonify({'error': 'No file or sheet selected'}), 400
    file = request.files['file']
    sheet = request.form['sheet']
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type'}), 400

    try:
        # Read headers from row 7 (index 6)
        df = pd.read_excel(file, sheet_name=sheet, header=6, nrows=1)
        columns = df.columns.tolist()
        default_columns = ['BLUEBOX IPAD STRUCTURE', 'BLUEBOX WOW STRUCTURE']
        preselected = [col for col in default_columns if col in columns]
        return jsonify({'columns': columns, 'preselected': preselected})
    except Exception as e:
        return jsonify({'error': f'Error reading sheet: {str(e)}'}), 500

@app.route('/check_filenames', methods=['POST'])
def check_filenames():
    if 'file' not in request.files or not request.form.get('sheet') or not request.form.get('columns'):
        return jsonify({'error': 'No file, sheet, or columns selected'}), 400
    file = request.files['file']
    sheet = request.form['sheet']
    columns = request.form.getlist('columns')
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type'}), 400

    try:
        # Read data from row 8 (header in row 7, index 6)
        df = pd.read_excel(file, sheet_name=sheet, header=6, dtype=str)
        df = df.fillna('')

        # Validation logic
        invalid_chars = r'[<>:"\\|?*]'
        reserved_names = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'LPT1']
        max_path_length = 260
        results = []
        duplicates = set()

        for idx, row in df.iterrows():
            for col in columns:
                filename = str(row.get(col, ''))
                if filename:
                    # Check invalid characters
                    import re
                    if re.search(invalid_chars, filename):
                        found_chars = re.findall(invalid_chars, filename)
                        suggested = re.sub(invalid_chars, '_', filename)
                        results.append({
                            'row': idx + 8,  # Adjust for Excel row (header in 7, data from 8)
                            'column': col,
                            'filename': filename,
                            'issue': f'Invalid characters: {", ".join(found_chars)}',
                            'suggested': suggested
                        })

                    # Check path length
                    if len(filename) > max_path_length:
                        results.append({
                            'row': idx + 8,
                            'column': col,
                            'filename': filename,
                            'issue': f'Filename exceeds {max_path_length} characters',
                            'suggested': filename[:max_path_length-10] + '...'
                        })

                    # Check reserved names
                    filename_part = filename.split('/')[-1].split('.')[0]
                    if filename_part.upper() in reserved_names:
                        results.append({
                            'row': idx + 8,
                            'column': col,
                            'filename': filename,
                            'issue': f'Uses reserved name: {filename_part}',
                            'suggested': f'_{filename_part}_'
                        })

                    # Check trailing spaces/periods
                    if filename.strip() != filename:
                        results.append({
                            'row': idx + 8,
                            'column': col,
                            'filename': filename,
                            'issue': 'Contains trailing spaces',
                            'suggested': filename.strip()
                        })
                    if filename.endswith('.'):
                        results.append({
                            'row': idx + 8,
                            'column': col,
                            'filename': filename,
                            'issue': 'Ends with a period',
                            'suggested': filename.rstrip('.')
                        })

                    # Check duplicates
                    lower_filename = filename.lower()
                    if lower_filename in duplicates:
                        results.append({
                            'row': idx + 8,
                            'column': col,
                            'filename': filename,
                            'issue': 'Duplicate filename (case-insensitive)',
                            'suggested': f'{filename_part}_copy{idx}.{filename.split(".")[-1]}'
                        })
                    else:
                        duplicates.add(lower_filename)

                    # Check double slashes
                    if '//' in filename:
                        results.append({
                            'row': idx + 8,
                            'column': col,
                            'filename': filename,
                            'issue': 'Contains double slashes',
                            'suggested': filename.replace('//', '/')
                        })

        return jsonify({'results': results})
    except Exception as e:
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

@app.route('/download_csv', methods=['POST'])
def download_csv():
    results = request.json.get('results', [])
    if not results:
        return jsonify({'error': 'No results to download'}), 400

    output = io.StringIO()
    output.write('Row,Column,Filename,Issue,Suggested Fix\n')
    for result in results:
        output.write(f'"{result["row"]}","{result["column"]}","{result["filename"].replace("\"", "\"\"")}","{result["issue"].replace("\"", "\"\"")}","{result["suggested"].replace("\"", "\"\"")}"\n')
    
    output.seek(0)
    return send_file(
        io.BytesIO(output.getvalue().encode('utf-8')),
        mimetype='text/csv',
        as_attachment=True,
        download_name='filename_qc_report.csv'
    )

if __name__ == '__main__':
    app.run(debug=True)
