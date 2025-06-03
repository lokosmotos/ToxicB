from flask import Flask, render_template, request, redirect, flash
import pandas as pd
import re
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'your-secret-key'  # Change this in production

# Upload folder
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET', 'POST'])
def index():
    results = []
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)

        if file and file.filename.endswith(('.xls', '.xlsx')):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(filepath)

            try:
                df = pd.read_excel(filepath)
            except Exception as e:
                flash(f"Failed to read Excel file: {e}")
                return redirect(request.url)

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
                            found = ', '.join(set(invalid_chars.findall(filename)))
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
        else:
            flash("Invalid file type. Please upload .xls or .xlsx files.")

    return render_template('index.html', results=results)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Required for Render
    app.run(host="0.0.0.0", port=port)
