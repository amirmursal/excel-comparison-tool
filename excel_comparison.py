#!/usr/bin/env python3
"""
Excel File Comparison Tool
Compare Patient Names between two Excel files and add status column
"""

from flask import Flask, render_template_string, request, jsonify, send_file, redirect
import pandas as pd
import os
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Global variables to store session data
raw_data = None
previous_data = None
raw_filename = None
previous_filename = None
comparison_result = None

# HTML Template for Excel Comparison
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel File Comparison Tool</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .container { 
            max-width: 1400px; 
            margin: 0 auto; 
            background: white; 
            border-radius: 15px; 
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            overflow: hidden;
        }
        .header { 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white; 
            padding: 30px; 
            text-align: center; 
        }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.2em; opacity: 0.9; }
        .content { padding: 30px; }
        .section { 
            margin: 25px 0; 
            padding: 25px; 
            border: 1px solid #e0e0e0; 
            border-radius: 10px; 
            background: #fafafa;
        }
        .section h3 { 
            color: #333; 
            margin-bottom: 20px; 
            font-size: 1.4em;
            border-bottom: 2px solid #667eea;
            padding-bottom: 10px;
        }
        .form-group { margin: 15px 0; }
        label { 
            display: block; 
            margin-bottom: 8px; 
            font-weight: 600; 
            color: #555;
        }
        input[type="file"] { 
            width: 100%; 
            padding: 12px; 
            border: 2px solid #ddd; 
            border-radius: 8px; 
            font-size: 16px;
            transition: border-color 0.3s;
        }
        input[type="file"]:focus { 
            outline: none; 
            border-color: #667eea; 
        }
        button { 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white; 
            padding: 12px 25px; 
            border: none; 
            border-radius: 8px; 
            cursor: pointer; 
            margin: 5px; 
            font-size: 16px;
            font-weight: 600;
            transition: transform 0.2s;
        }
        button:hover { 
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }
        .file-status { 
            background: #f8f9fa; 
            padding: 15px; 
            border-radius: 8px; 
            margin: 15px 0; 
            border-left: 4px solid #667eea;
        }
        .status-success { 
            background: #d4edda; 
            color: #155724; 
            border-color: #c3e6cb; 
        }
        .status-info { 
            background: #d1ecf1; 
            color: #0c5460; 
            border-color: #bee5eb; 
        }
        
        .status-message {
            background: #f3e5f5;
            border: 2px solid #9c27b0;
            color: #4a148c;
            padding: 20px;
            border-radius: 10px;
            margin: 15px 0;
            font-size: 16px;
            line-height: 1.6;
            white-space: pre-line;
            box-shadow: 0 2px 8px rgba(156, 39, 176, 0.2);
        }
        
        .reset-btn {
            background: linear-gradient(135deg, #ff6b6b, #ee5a52);
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 16px;
            font-weight: 600;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(255, 107, 107, 0.3);
        }
        
        .reset-btn:hover {
            background: linear-gradient(135deg, #ff5252, #d32f2f);
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(255, 107, 107, 0.4);
        }
        
        .reset-btn:active {
            transform: translateY(0);
            box-shadow: 0 2px 10px rgba(255, 107, 107, 0.3);
        }
        .comparison-results { 
            background: #1e1e1e; 
            color: #f8f8f2; 
            padding: 20px; 
            border-radius: 8px; 
            white-space: pre-wrap; 
            font-family: 'Courier New', monospace; 
            max-height: 400px; 
            overflow-y: auto;
            border: 1px solid #333;
        }
        .two-column { 
            display: grid; 
            grid-template-columns: 1fr 1fr; 
            gap: 20px; 
        }
        .loading { 
            display: none; 
            text-align: center; 
            padding: 20px; 
        }
        .spinner { 
            border: 4px solid #f3f3f3; 
            border-top: 4px solid #667eea; 
            border-radius: 50%; 
            width: 40px; 
            height: 40px; 
            animation: spin 1s linear infinite; 
            margin: 0 auto;
        }
        @keyframes spin { 
            0% { transform: rotate(0deg); } 
            100% { transform: rotate(360deg); } 
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Excel File Comparison Tool</h1>
            <p>Compare Patient Names between two Excel files and add status column</p>
        </div>

        <div class="content">

            <!-- File Upload Section -->
            <div class="section">
                <h3>📁 Upload Excel Files</h3>
                <div class="two-column">
                    <div>
                        <h4>Raw Excel File</h4>
                        <form action="/upload_raw" method="post" enctype="multipart/form-data" id="raw-form">
                            <div class="form-group">
                                <label for="raw_file">Select Raw Excel File:</label>
                                <input type="file" id="raw_file" name="file" accept=".xlsx,.xls" required>
                            </div>
                            <button type="submit" id="raw-btn">📤 Upload Raw File</button>
                        </form>
                        <div class="loading" id="raw-loading">
                            <div class="spinner"></div>
                            <p>Processing raw file...</p>
                        </div>
                    </div>
                    
                    <div>
                        <h4>Previous Excel File</h4>
                        <form action="/upload_previous" method="post" enctype="multipart/form-data" id="previous-form">
                            <div class="form-group">
                                <label for="previous_file">Select Previous Excel File:</label>
                                <input type="file" id="previous_file" name="file" accept=".xlsx,.xls" required>
                            </div>
                            <button type="submit" id="previous-btn">📤 Upload Previous File</button>
                        </form>
                        <div class="loading" id="previous-loading">
                            <div class="spinner"></div>
                            <p>Processing previous file...</p>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Status Messages -->
            {% if comparison_result %}
            <div class="section">
                <h3>📢 Status Messages</h3>
                <div class="status-message">
                    {{ comparison_result | safe }}
                </div>
            </div>
            {% endif %}

            <!-- File Status -->
            <div class="section">
                <h3>📊 File Status</h3>
                <div class="file-status">
                    {% if raw_filename %}
                        <div class="status-success">
                            ✅ Raw File: {{ raw_filename }}<br>
                            📋 Sheets: {{ raw_data.keys() | list | length if raw_data else 0 }}
                        </div>
                    {% else %}
                        <div class="status-info">
                            ℹ️ No raw file uploaded yet.
                        </div>
                    {% endif %}
                    
                    {% if previous_filename %}
                        <div class="status-success">
                            ✅ Previous File: {{ previous_filename }}<br>
                            📋 Sheets: {{ previous_data.keys() | list | length if previous_data else 0 }}
                        </div>
                    {% else %}
                        <div class="status-info">
                            ℹ️ No previous file uploaded yet.
                        </div>
                    {% endif %}
                </div>
            </div>

            <!-- Comparison Section -->
            {% if raw_data and previous_data %}
            <div class="section">
                <h3>🔄 Compare Files</h3>
                <form action="/compare" method="post" id="compare-form">
                    <div class="form-group">
                        <label for="raw_sheet">Select Raw File Sheet:</label>
                        <select id="raw_sheet" name="raw_sheet" required style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                            {% for sheet_name in raw_data.keys() %}
                                <option value="{{ sheet_name }}">{{ sheet_name }}</option>
                            {% endfor %}
                        </select>
                    </div>
                    <div class="form-group">
                        <label for="previous_sheet">Select Previous File Sheet:</label>
                        <select id="previous_sheet" name="previous_sheet" required style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                            {% for sheet_name in previous_data.keys() %}
                                <option value="{{ sheet_name }}">{{ sheet_name }}</option>
                            {% endfor %}
                        </select>
                    </div>
                    <button type="submit" id="compare-btn">🔄 Compare Patient Names</button>
                </form>
                <div class="loading" id="compare-loading">
                    <div class="spinner"></div>
                    <p>Comparing files...</p>
                </div>
            </div>
            {% endif %}

            <!-- Results Section -->
            <div class="section">
                <h3>📝 Comparison Results</h3>
                <div class="comparison-results" id="results">
                    {% if comparison_result %}
                        {{ comparison_result }}
                    {% else %}
                        Upload both files and click "Compare Patient Names" to see results...
                    {% endif %}
                </div>
            </div>

            <!-- Download Section -->
            {% if comparison_result and 'Status column added' in comparison_result %}
            <div class="section">
                <h3>💾 Download Results</h3>
                <form action="/download_result" method="post">
                    <div class="form-group">
                        <label for="output_filename">Output filename (optional):</label>
                        <input type="text" id="output_filename" name="filename" 
                               placeholder="comparison_result.xlsx" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                    </div>
                    <button type="submit">💾 Download Result File</button>
                </form>
            </div>
            {% endif %}

            <!-- Reset Section -->
            <div class="section">
                <h3>🔄 Reset Application</h3>
                <p>Clear all uploaded files and reset the application to start fresh.</p>
                <form action="/reset_app" method="post" onsubmit="return confirm('Are you sure you want to reset the application? This will clear all uploaded files and data.')">
                    <button type="submit" class="reset-btn">🗑️ Reset Application</button>
                </form>
            </div>
        </div>
    </div>

    <script>
        // Form submission with loading states
        document.getElementById('raw-form').addEventListener('submit', function() {
            document.getElementById('raw-loading').style.display = 'block';
            document.getElementById('raw-btn').disabled = true;
        });

        document.getElementById('previous-form').addEventListener('submit', function() {
            document.getElementById('previous-loading').style.display = 'block';
            document.getElementById('previous-btn').disabled = true;
        });

        document.getElementById('compare-form').addEventListener('submit', function() {
            document.getElementById('compare-loading').style.display = 'block';
            document.getElementById('compare-btn').disabled = true;
        });

        // Auto-scroll results to bottom
        function scrollResults() {
            const results = document.getElementById('results');
            results.scrollTop = results.scrollHeight;
        }

        // Scroll results on page load
        window.onload = function() {
            scrollResults();
        }
    </script>
</body>
</html>
"""

def compare_patient_names(raw_df, previous_df):
    """Compare Patient Names between two DataFrames and add Status column"""
    try:
        # Debug: Show available columns
        raw_columns = list(raw_df.columns)
        previous_columns = list(previous_df.columns)
        
        # Function to find patient name column (case-insensitive, handles spaces)
        def find_patient_name_column(columns):
            for col in columns:
                if 'patient' in col.lower() and 'name' in col.lower():
                    return col
            return None
        
        # Find patient name columns
        raw_patient_col = find_patient_name_column(raw_columns)
        previous_patient_col = find_patient_name_column(previous_columns)
        
        # Check if Patient Name column exists in both DataFrames
        if not raw_patient_col:
            return f"❌ Error: 'Patient Name' column not found in raw file.\nAvailable columns: {raw_columns}", None
        
        if not previous_patient_col:
            return f"❌ Error: 'Patient Name' column not found in previous file.\nAvailable columns: {previous_columns}", None
        
        # Get unique patient names from previous file
        previous_patients = set(previous_df[previous_patient_col].dropna().astype(str).str.strip())
        
        # Create a copy of raw data to avoid modifying original
        result_df = raw_df.copy()
        
        # Find the position of the Patient Name column
        patient_col_index = result_df.columns.get_loc(raw_patient_col)
        
        # Insert Status column right after Patient Name column
        result_df.insert(patient_col_index + 1, 'Status', '')
        
        # Compare and set status
        for idx, patient_name in enumerate(result_df[raw_patient_col]):
            if pd.notna(patient_name):
                patient_str = str(patient_name).strip()
                if patient_str in previous_patients:
                    result_df.at[idx, 'Status'] = 'Done'
        
        # Count statistics
        total_patients = len(result_df)
        done_patients = (result_df['Status'] == 'Done').sum()
        not_done_patients = total_patients - done_patients
        
        # Generate result message
        result_message = f"""✅ Comparison completed successfully!

📊 Statistics:
- Total patients in raw file: {total_patients}
- Patients found in previous file: {done_patients}
- Patients not found in previous file: {not_done_patients}
- Completion rate: {(done_patients/total_patients*100):.1f}%

📋 Columns used:
- Raw file: "{raw_patient_col}"
- Previous file: "{previous_patient_col}"

📋 Status column added to raw file with:
- "Done" for patients found in previous file
- Empty cell for patients not found in previous file

💾 Ready to download the result file!"""
        
        return result_message, result_df
        
    except Exception as e:
        return f"❌ Error during comparison: {str(e)}", None

@app.route('/')
def root():
    return redirect('/comparison')

@app.route('/comparison')
def comparison_index():
    global raw_data, previous_data, raw_filename, previous_filename, comparison_result
    return render_template_string(HTML_TEMPLATE, 
                                raw_data=raw_data, 
                                previous_data=previous_data,
                                raw_filename=raw_filename,
                                previous_filename=previous_filename,
                                comparison_result=comparison_result)

@app.route('/upload_raw', methods=['POST'])
def upload_raw_file():
    global raw_data, raw_filename, comparison_result
    
    if 'file' not in request.files:
        comparison_result = "❌ Error: No file provided"
        return redirect('/comparison')
    
    file = request.files['file']
    if file.filename == '':
        comparison_result = "❌ Error: No file selected"
        return redirect('/comparison')
    
    try:
        # Save uploaded file
        filename = secure_filename(file.filename)
        file.save(filename)
        
        # Load Excel file
        raw_data = pd.read_excel(filename, sheet_name=None)
        raw_filename = filename
        
        comparison_result = f"✅ Raw file uploaded successfully! Loaded {len(raw_data)} sheets: {', '.join(list(raw_data.keys()))}"
        return redirect('/comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error uploading raw file: {str(e)}"
        return redirect('/comparison')

@app.route('/upload_previous', methods=['POST'])
def upload_previous_file():
    global previous_data, previous_filename, comparison_result
    
    if 'file' not in request.files:
        comparison_result = "❌ Error: No file provided"
        return redirect('/comparison')
    
    file = request.files['file']
    if file.filename == '':
        comparison_result = "❌ Error: No file selected"
        return redirect('/comparison')
    
    try:
        # Save uploaded file
        filename = secure_filename(file.filename)
        file.save(filename)
        
        # Load Excel file
        previous_data = pd.read_excel(filename, sheet_name=None)
        previous_filename = filename
        
        comparison_result = f"✅ Previous file uploaded successfully! Loaded {len(previous_data)} sheets: {', '.join(list(previous_data.keys()))}"
        return redirect('/comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error uploading previous file: {str(e)}"
        return redirect('/comparison')

@app.route('/compare', methods=['POST'])
def compare_files():
    global raw_data, previous_data, comparison_result
    
    if not raw_data or not previous_data:
        comparison_result = "❌ Error: Please upload both files first"
        return redirect('/comparison')
    
    raw_sheet = request.form.get('raw_sheet')
    previous_sheet = request.form.get('previous_sheet')
    
    if not raw_sheet or not previous_sheet:
        comparison_result = "❌ Error: Please select sheets for both files"
        return redirect('/comparison')
    
    try:
        # Get the selected sheets
        raw_df = raw_data[raw_sheet]
        previous_df = previous_data[previous_sheet]
        
        # Perform comparison
        result_message, result_df = compare_patient_names(raw_df, previous_df)
        
        if result_df is not None:
            # Store the result for download
            comparison_result = result_message
            # Update the raw_data with the result
            raw_data[raw_sheet] = result_df
        else:
            comparison_result = result_message
        
        return redirect('/comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error comparing files: {str(e)}"
        return redirect('/comparison')

@app.route('/download_result', methods=['POST'])
def download_result():
    global raw_data, raw_filename
    
    if not raw_data:
        return jsonify({'error': 'No data to download'}), 400
    
    filename = request.form.get('filename', '').strip()
    if not filename:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"comparison_result_{timestamp}.xlsx"
    
    try:
        # Create a temporary file
        import tempfile
        temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
        
        try:
            with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                for sheet_name, df in raw_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            return send_file(temp_path, as_attachment=True, download_name=filename)
            
        finally:
            # Clean up temporary file
            os.close(temp_fd)
            if os.path.exists(temp_path):
                os.unlink(temp_path)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/reset_app', methods=['POST'])
def reset_app():
    global raw_data, previous_data, raw_filename, previous_filename, comparison_result
    
    try:
        # Reset all global variables
        raw_data = {}
        previous_data = {}
        raw_filename = None
        previous_filename = None
        comparison_result = "🔄 Application reset successfully! All files and data have been cleared."
        
        return redirect('/comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error resetting application: {str(e)}"
        return redirect('/comparison')

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5002))
    debug = os.environ.get('FLASK_ENV') == 'development'
    
    print("🚀 Starting Excel Comparison Tool...")
    print(f"📱 Open your browser and go to: http://localhost:{port}")
    app.run(debug=debug, host='0.0.0.0', port=port)
