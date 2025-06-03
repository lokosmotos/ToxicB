from flask import Flask, render_template, request, redirect, flash
import pandas as pd
import re
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'your-secret-key'
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET', 'POST'])
def index():
    results = []
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith(('.xls', '.xlsx')):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(filepath)

            df = pd.read_excel(filepath)
            filename_columns = ['RAVE CENTRIC FILENAME', 'BLUEBOX WOW FILENAME']
            invalid_chars = re.compile(r'[<>:"\\|?*]')
            max_path_length = 260
            reserved_names = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'LPT1']

            for idx, row in df.iterrows():
                for column in filename_columns:
                    filename = row.get(column)
                    if pd.notna(filename) and isinstance(filename, str):
                        issues = []
                        if invalid_chars.search(filename):
                            found = ', '.join(invalid_chars.findall(filename))
                            issues.append(f"Invalid characters: {found}")
                        if len(filename) > max_path_length:
                            issues.append("Exceeds 260 characters.")
                        if filename.strip() != filename:
                            issues.append("Has leading/trailing spaces.")
                        if filename.endswith('.'):
                            issues.append("Ends with a period.")
                        name_part = filename.split('/')[-1].split('.')[0]
                        if name_part.upper() in reserved_names:
                            issues.append(f"Reserved name: {name_part}")
                        if issues:
                            results.append(f"Row {idx + 2}, Column '{column}': {', '.join(issues)}")

    return render_template('index.html', results=results)

