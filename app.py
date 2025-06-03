from flask import Flask, render_template, request
import pandas as pd
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/filename-checker', methods=['GET', 'POST'])
def filename_checker():
    results = []
    if request.method == 'POST':
        uploaded_file = request.files['file']
        if uploaded_file and uploaded_file.filename.endswith(('.xls', '.xlsx')):
            try:
                df = pd.read_excel(uploaded_file)

                if 'Filename' not in df.columns:
                    results.append({'filename': '', 'status': 'Missing "Filename" column'})
                else:
                    for filename in df['Filename']:
                        if pd.isna(filename):
                            status = 'Empty filename'
                        elif not str(filename).endswith(('.xlsx', '.xls')):
                            status = 'Invalid extension'
                        elif ' ' in str(filename):
                            status = 'Contains spaces'
                        else:
                            status = 'OK'
                        results.append({'filename': filename, 'status': status})
            except Exception as e:
                results.append({'filename': '', 'status': f'Error: {str(e)}'})
        else:
            results.append({'filename': '', 'status': 'Unsupported file type. Upload .xls or .xlsx only.'})
    return render_template('filename_qc.html', results=results)

if __name__ == '__main__':
    app.run(debug=True)
