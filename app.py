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

# Configuration from environment variables
app.secret_key = os.environ.get('SECRET_KEY', 'your-secret-key-here-change-in-production')
app.config['FLASK_ENV'] = os.environ.get('FLASK_ENV', 'development')
app.config['DEBUG'] = os.environ.get('FLASK_DEBUG', 'True').lower() == 'true'

# Session configuration
app.config['SESSION_TYPE'] = 'filesystem'
app.config['PERMANENT_SESSION_LIFETIME'] = 3600  # 1 hour

# Configuration
UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER', 'uploads')
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_FILE_SIZE = int(os.environ.get('MAX_FILE_SIZE', 16 * 1024 * 1024))  # 16MB default

# Create upload folder if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Global storage for file paths (session alternative)
file_storage = {}

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
        values = rule_data['value1'] if isinstance(rule_data['value1'], list) else [rule_data['value1']]
        values_str = '_'.join(values)
        return f"{col}_{values_str}.xlsx"
    elif rule_type == 'and':
        # Start with first two columns
        col1 = rule_data['column1']
        values1 = rule_data['value1'] if isinstance(rule_data['value1'], list) else [rule_data['value1']]
        col2 = rule_data['column2']
        values2 = rule_data['value2'] if isinstance(rule_data['value2'], list) else [rule_data['value2']]
        values1_str = '_'.join(values1)
        values2_str = '_'.join(values2)
        filename = f"{col1}_{values1_str}_{col2}_{values2_str}"
        
        # Add additional columns
        additional_columns = rule_data.get('additional_columns', [])
        additional_values = rule_data.get('additional_values', [])
        
        for i, col in enumerate(additional_columns):
            if i < len(additional_values):
                values = additional_values[i] if isinstance(additional_values[i], list) else [additional_values[i]]
                values_str = '_'.join(values)
                filename += f"_{col}_{values_str}"
        
        return f"{filename}.xlsx"
    elif rule_type == 'or':
        # Start with first two columns
        col1 = rule_data['column1']
        values1 = rule_data['value1'] if isinstance(rule_data['value1'], list) else [rule_data['value1']]
        col2 = rule_data['column2']
        values2 = rule_data['value2'] if isinstance(rule_data['value2'], list) else [rule_data['value2']]
        values1_str = '_'.join(values1)
        values2_str = '_'.join(values2)
        filename = f"{col1}_{values1_str}_OR_{col2}_{values2_str}"
        
        # Add additional columns
        additional_columns = rule_data.get('additional_columns', [])
        additional_values = rule_data.get('additional_values', [])
        
        for i, col in enumerate(additional_columns):
            if i < len(additional_values):
                values = additional_values[i] if isinstance(additional_values[i], list) else [additional_values[i]]
                values_str = '_'.join(values)
                filename += f"_OR_{col}_{values_str}"
        
        return f"{filename}.xlsx"
    
    return f"split_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

def apply_rule(df, rule_data):
    """Apply rule to DataFrame and return filtered data"""
    rule_type = rule_data['rule_type']
    
    if rule_type == 'single':
        column = rule_data['column1']
        values = rule_data['value1'] if isinstance(rule_data['value1'], list) else [rule_data['value1']]
        return df[df[column].isin(values)]
    
    elif rule_type == 'and':
        # Start with first two columns
        col1 = rule_data['column1']
        values1 = rule_data['value1'] if isinstance(rule_data['value1'], list) else [rule_data['value1']]
        col2 = rule_data['column2']
        values2 = rule_data['value2'] if isinstance(rule_data['value2'], list) else [rule_data['value2']]
        
        # Apply first two conditions
        filtered_df = df[(df[col1].isin(values1)) & (df[col2].isin(values2))]
        
        # Apply additional columns (3-6) if they exist
        additional_columns = rule_data.get('additional_columns', [])
        additional_values = rule_data.get('additional_values', [])
        
        for i, col in enumerate(additional_columns):
            if i < len(additional_values):
                values = additional_values[i] if isinstance(additional_values[i], list) else [additional_values[i]]
                filtered_df = filtered_df[filtered_df[col].isin(values)]
        
        return filtered_df
    
    elif rule_type == 'or':
        # Start with first two columns
        col1 = rule_data['column1']
        values1 = rule_data['value1'] if isinstance(rule_data['value1'], list) else [rule_data['value1']]
        col2 = rule_data['column2']
        values2 = rule_data['value2'] if isinstance(rule_data['value2'], list) else [rule_data['value2']]
        
        # Apply first two conditions
        filtered_df = df[(df[col1].isin(values1)) | (df[col2].isin(values2))]
        
        # Apply additional columns (3-6) if they exist
        additional_columns = rule_data.get('additional_columns', [])
        additional_values = rule_data.get('additional_values', [])
        
        for i, col in enumerate(additional_columns):
            if i < len(additional_values):
                values = additional_values[i] if isinstance(additional_values[i], list) else [additional_values[i]]
                filtered_df = filtered_df | df[df[col].isin(values)]
        
        return filtered_df
    
    return df

@app.route('/test-session')
def test_session():
    """Test session functionality"""
    print(f"Current session: {dict(session)}")  # Debug log
    return jsonify({
        'session_data': dict(session),
        'file_path': session.get('file_path'),
        'file_exists': os.path.exists(session.get('file_path', '')) if session.get('file_path') else False
    })

@app.route('/')
def index():
    """Main page"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and return column data"""
    try:
        print("=== UPLOAD FILE CALLED ===")  # Debug log
        
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
            return jsonify({'error': f'File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.'}), 400
        
        # Save file temporarily
        filename = secure_filename(file.filename)
        session_id = str(uuid.uuid4())
        file_path = os.path.join(UPLOAD_FOLDER, f"{session_id}_{filename}")
        file.save(file_path)
        
        print(f"File saved to: {file_path}")  # Debug log
        
        # Read Excel file
        df = pd.read_excel(file_path)
        print(f"DataFrame shape: {df.shape}")  # Debug log
        
        # Get column information
        columns = df.columns.tolist()
        column_values = {}
        
        for col in columns:
            unique_values = df[col].dropna().unique().tolist()
            # Convert to strings and limit to first 50 unique values
            column_values[col] = [str(val) for val in unique_values[:50]]
        
        # Store file path in global storage and session
        file_storage[session_id] = file_path
        session['session_id'] = session_id
        session['file_path'] = file_path
        session['columns'] = columns
        session.permanent = True  # Make session permanent
        
        print(f"Session ID: {session_id}")  # Debug log
        print(f"File stored in global storage: {file_storage.get(session_id)}")  # Debug log
        
        return jsonify({
            'success': True,
            'columns': columns,
            'column_values': column_values,
            'total_rows': len(df),
            'session_id': session_id
        })
        
    except Exception as e:
        print(f"Error in upload_file: {str(e)}")  # Debug log
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

@app.route('/process', methods=['POST'])
def process_rules():
    """Process rules and generate Excel files"""
    try:
        print("=== PROCESS RULES CALLED ===")  # Debug log
        data = request.get_json()
        print(f"Received data: {data}")  # Debug log
        rules = data.get('rules', [])
        session_id = data.get('session_id')  # Get session ID from request
        print(f"Rules: {rules}")  # Debug log
        print(f"Session ID from request: {session_id}")  # Debug log
        
        if not rules:
            return jsonify({'error': 'No rules provided'}), 400
        
        # Get file path from global storage using session ID
        file_path = file_storage.get(session_id) if session_id else None
        
        print(f"File path from global storage: {file_path}")  # Debug log
        print(f"Available session IDs: {list(file_storage.keys())}")  # Debug log
        
        if not file_path or not os.path.exists(file_path):
            print(f"File exists: {os.path.exists(file_path) if file_path else False}")  # Debug log
            return jsonify({'error': 'No file uploaded or file not found'}), 400
        
        # Read the Excel file
        df = pd.read_excel(file_path)
        print(f"DataFrame shape: {df.shape}")  # Debug log
        
        generated_files = []
        
        print(f"Starting to process {len(rules)} rules...")  # Debug log
        
        for i, rule in enumerate(rules):
            print(f"Processing rule {i + 1}/{len(rules)}: {rule}")  # Debug log
            try:
                # Apply rule to get filtered data
                filtered_df = apply_rule(df, rule)
                print(f"Rule {i + 1} filtered data shape: {filtered_df.shape}")  # Debug log
                
                # Skip if no data matches the rule
                if len(filtered_df) == 0:
                    print(f"Rule {i + 1}: No data matches rule, skipping")  # Debug log
                    continue
                
                # Generate filename
                filename = generate_filename(rule)
                print(f"Rule {i + 1} generated filename: {filename}")  # Debug log
                
                # Save filtered data to new Excel file
                output_path = os.path.join(UPLOAD_FOLDER, filename)
                filtered_df.to_excel(output_path, index=False)
                print(f"Rule {i + 1} saved file to: {output_path}")  # Debug log
                
                generated_files.append({
                    'filename': filename,
                    'rows': len(filtered_df),
                    'download_url': f'/download/{filename}'
                })
                print(f"Rule {i + 1} added to generated_files. Total so far: {len(generated_files)}")  # Debug log
                
            except Exception as e:
                print(f"Error processing rule {i + 1}: {str(e)}")  # Debug log
                continue
        
        print(f"Final result: Generated {len(generated_files)} files out of {len(rules)} rules")  # Debug log
        print(f"Generated files: {generated_files}")  # Debug log
        return jsonify({
            'success': True,
            'files': generated_files,
            'total_files': len(generated_files)
        })
        
    except Exception as e:
        print(f"Error in process_rules: {str(e)}")  # Debug log
        return jsonify({'error': f'Error processing rules: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """Download generated Excel file"""
    try:
        print(f"=== DOWNLOAD REQUESTED ===")  # Debug log
        print(f"Filename: {filename}")  # Debug log
        
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        print(f"File path: {file_path}")  # Debug log
        print(f"File exists: {os.path.exists(file_path)}")  # Debug log
        
        if os.path.exists(file_path):
            print(f"File size: {os.path.getsize(file_path)} bytes")  # Debug log
            return send_file(
                file_path, 
                as_attachment=True, 
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            print(f"File not found: {file_path}")  # Debug log
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        print(f"Error in download_file: {str(e)}")  # Debug log
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
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=app.config['DEBUG'], host='0.0.0.0', port=port) 