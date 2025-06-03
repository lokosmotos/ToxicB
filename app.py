from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import io
import os
import csv
from werkzeug.utils import secure_filename
from datetime import datetime
import json

app = Flask(__name__)

# Configuration
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}
app.config['ANALYTICS_FILE'] = 'usage_analytics.json'

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def update_analytics(action, filename=None, sheet=None, issues_found=0):
    """Track usage statistics for the dashboard"""
    try:
        try:
            with open(app.config['ANALYTICS_FILE'], 'r') as f:
                data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            data = {
                'total_checks': 0,
                'files_processed': 0,
                'issues_found': 0,
                'recent_files': [],
                'daily_stats': {}
            }
        
        today = datetime.now().strftime('%Y-%m-%d')
        
        # Update general stats
        data['total_checks'] += 1
        if action == 'file_processed':
            data['files_processed'] += 1
            data['issues_found'] += issues_found
            
            # Add to recent files (limit to last 5)
            recent = {
                'filename': filename,
                'date': datetime.now().strftime('%Y-%m-%d %H:%M'),
                'sheet': sheet,
                'issues': issues_found
            }
            data['recent_files'].insert(0, recent)
            data['recent_files'] = data['recent_files'][:5]
            
            # Update daily stats
            if today not in data['daily_stats']:
                data['daily_stats'][today] = {
                    'files': 0,
                    'issues': 0
                }
            data['daily_stats'][today]['files'] += 1
            data['daily_stats'][today]['issues'] += issues_found
        
        with open(app.config['ANALYTICS_FILE'], 'w') as f:
            json.dump(data, f)
    except Exception as e:
        print(f"Error updating analytics: {str(e)}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/dashboard')
def dashboard():
    """Dashboard showing usage statistics"""
    try:
        with open(app.config['ANALYTICS_FILE'], 'r') as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {
            'total_checks': 0,
            'files_processed': 0,
            'issues_found': 0,
            'recent_files': [],
            'daily_stats': {}
        }
    
    # Prepare data for charts
    dates = sorted(data['daily_stats'].keys())[-7:]  # Show last 7 days
    files_data = [data['daily_stats'].get(d, {}).get('files', 0) for d in dates]
    issues_data = [data['daily_stats'].get(d, {}).get('issues', 0) for d in dates]
    
    return render_template('dashboard.html', 
                         total_checks=data['total_checks'],
                         files_processed=data['files_processed'],
                         issues_found=data['issues_found'],
                         recent_files=data['recent_files'],
                         chart_data={
                             'dates': dates,
                             'files': files_data,
                             'issues': issues_data
                         })

# [Keep all your existing routes: /get_sheets, /get_columns, /check_filenames, /download_csv]

# Modify the check_filenames route to update analytics
@app.route('/check_filenames', methods=['POST'])
def check_filenames():
    # [Keep all your existing check_filenames code until the end]
    
    # Before returning results, update analytics
    update_analytics('file_processed', 
                    filename=file.filename,
                    sheet=sheet,
                    issues_found=len(results))
    
    return jsonify({'results': results})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
