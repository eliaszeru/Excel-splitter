"""
Excel Splitter Web Application
Splits master Excel files into multiple files based on user-defined rules
"""
import os
import pandas as pd
import json
import uuid
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
import tempfile
import shutil

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Change this in production

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_FILE_SIZE = 16 * 1024 * 1024  # 16MB

# Create upload folder if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def generate_filename(rule_data):
    """Generate filename based on rule data"""
    if rule_data.get('custom_name'):
        return f"{rule_data['custom_name']}.xlsx"
    
    # Auto-generate name based on rule
    rule_type = rule_data['rule_type']
    if rule_type == 'single':
        col = rule_data['column1']
        val = rule_data['value1']
        return f"{col}_{val}.xlsx"
    elif rule_type == 'and':
        col1 = rule_data['column1']
        val1 = rule_data['value1']
        col2 = rule_data['column2']
        val2 = rule_data['value2']
        return f"{col1}_{val1}_{col2}_{val2}.xlsx"
    elif rule_type == 'or':
        col1 = rule_data['column1']
        val1 = rule_data['value1']
        col2 = rule_data['column2']
        val2 = rule_data['value2']
        return f"{col1}_{val1}_OR_{col2}_{val2}.xlsx"
    
    return f"split_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

def apply_rule(df, rule_data):
    """Apply rule to DataFrame and return filtered data"""
    rule_type = rule_data['rule_type']
    
    if rule_type == 'single':
        column = rule_data['column1']
        value = rule_data['value1']
        return df[df[column] == value]
    
    elif rule_type == 'and':
        col1 = rule_data['column1']
        val1 = rule_data['value1']
        col2 = rule_data['column2']
        val2 = rule_data['value2']
        return df[(df[col1] == val1) & (df[col2] == val2)]
    
    elif rule_type == 'or':
        col1 = rule_data['column1']
        val1 = rule_data['value1']
        col2 = rule_data['column2']
        val2 = rule_data['value2']
        return df[(df[col1] == val1) | (df[col2] == val2)]
    
    return df

@app.route('/')
def index():
    """Main page"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and return column data"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Please upload Excel files only.'}), 400
        
        # Check file size
        file.seek(0, 2)  # Seek to end
        file_size = file.tell()
        file.seek(0)  # Reset to beginning
        
        if file_size > MAX_FILE_SIZE:
            return jsonify({'error': 'File too large. Maximum size is 16MB.'}), 400
        
        # Save file temporarily
        filename = secure_filename(file.filename)
        session_id = str(uuid.uuid4())
        file_path = os.path.join(UPLOAD_FOLDER, f"{session_id}_{filename}")
        file.save(file_path)
        
        # Read Excel file
        df = pd.read_excel(file_path)
        
        # Get column information
        columns = df.columns.tolist()
        column_values = {}
        
        for col in columns:
            unique_values = df[col].dropna().unique().tolist()
            # Convert to strings and limit to first 50 unique values
            column_values[col] = [str(val) for val in unique_values[:50]]
        
        # Store file path in session
        session['file_path'] = file_path
        session['columns'] = columns
        
        return jsonify({
            'success': True,
            'columns': columns,
            'column_values': column_values,
            'total_rows': len(df)
        })
        
    except Exception as e:
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

@app.route('/process', methods=['POST'])
def process_rules():
    """Process rules and generate Excel files"""
    try:
        data = request.get_json()
        rules = data.get('rules', [])
        
        if not rules:
            return jsonify({'error': 'No rules provided'}), 400
        
        file_path = session.get('file_path')
        if not file_path or not os.path.exists(file_path):
            return jsonify({'error': 'No file uploaded or file not found'}), 400
        
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        generated_files = []
        
        for rule in rules:
            # Apply rule to get filtered data
            filtered_df = apply_rule(df, rule)
            
            # Skip if no data matches the rule
            if len(filtered_df) == 0:
                continue
            
            # Generate filename
            filename = generate_filename(rule)
            
            # Save filtered data to new Excel file
            output_path = os.path.join(UPLOAD_FOLDER, filename)
            filtered_df.to_excel(output_path, index=False)
            
            generated_files.append({
                'filename': filename,
                'rows': len(filtered_df),
                'download_url': f'/download/{filename}'
            })
        
        return jsonify({
            'success': True,
            'files': generated_files,
            'total_files': len(generated_files)
        })
        
    except Exception as e:
        return jsonify({'error': f'Error processing rules: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """Download generated Excel file"""
    try:
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': f'Error downloading file: {str(e)}'}), 500

@app.route('/cleanup', methods=['POST'])
def cleanup_files():
    """Clean up uploaded and generated files"""
    try:
        file_path = session.get('file_path')
        if file_path and os.path.exists(file_path):
            os.remove(file_path)
        
        # Clean up generated files (older than 1 hour)
        current_time = datetime.now()
        for filename in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.isfile(file_path):
                file_time = datetime.fromtimestamp(os.path.getctime(file_path))
                if (current_time - file_time).total_seconds() > 3600:  # 1 hour
                    os.remove(file_path)
        
        session.clear()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': f'Error cleaning up files: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 