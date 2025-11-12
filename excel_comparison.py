#!/usr/bin/env python3
"""
Excel File Comparison Tool
Compare Patient IDs between two Excel files and add insurance columns
"""

from flask import Flask, render_template_string, request, jsonify, send_file, redirect
import pandas as pd
import os
from datetime import datetime
from werkzeug.utils import secure_filename
import re

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
            <p>Compare Patient ID (Appointment report) with PATID (Smart Assist) and add insurance columns</p>
        </div>

        <div class="content">

            <!-- File Upload Section -->
            <div class="section">
                <h3>📁 Upload Excel Files</h3>
                <div class="two-column">
                    <div>
                        <h4>Appointment Report File</h4>
                        <form action="/upload_raw" method="post" enctype="multipart/form-data" id="raw-form">
                            <div class="form-group">
                                <label for="raw_file">Select Appointment Report Excel File:</label>
                                <input type="file" id="raw_file" name="file" accept=".xlsx,.xls" required>
                            </div>
                            <button type="submit" id="raw-btn">📤 Upload Appointment Report</button>
                        </form>
                        <div class="loading" id="raw-loading">
                            <div class="spinner"></div>
                            <p>Processing Appointment Report...</p>
                        </div>
                    </div>
                    
                    <div>
                        <h4>Smart Assist File</h4>
                        <form action="/upload_previous" method="post" enctype="multipart/form-data" id="previous-form">
                            <div class="form-group">
                                <label for="previous_file">Select Smart Assist Excel File:</label>
                                <input type="file" id="previous_file" name="file" accept=".xlsx,.xls" required>
                            </div>
                            <button type="submit" id="previous-btn">📤 Upload Smart Assist File</button>
                        </form>
                        <div class="loading" id="previous-loading">
                            <div class="spinner"></div>
                            <p>Processing Smart Assist file...</p>
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
                            ✅ Appointment Report: {{ raw_filename }}<br>
                            📋 Sheets: {{ raw_data.keys() | list | length if raw_data else 0 }}
                        </div>
                    {% else %}
                        <div class="status-info">
                            ℹ️ No Appointment Report file uploaded yet.
                        </div>
                    {% endif %}
                    
                    {% if previous_filename %}
                        <div class="status-success">
                            ✅ Smart Assist File: {{ previous_filename }}<br>
                            📋 Sheets: {{ previous_data.keys() | list | length if previous_data else 0 }}
                        </div>
                    {% else %}
                        <div class="status-info">
                            ℹ️ No Smart Assist file uploaded yet.
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
                    <button type="submit" id="compare-btn">🔄 Compare and Add Insurance Columns</button>
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
                        Upload both files and click "Compare and Add Insurance Columns" to see results...
                    {% endif %}
                </div>
            </div>

            <!-- Download Section -->
            {% if comparison_result and 'Comparison completed successfully' in comparison_result %}
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

# State abbreviations mapping
STATE_ABBREVIATIONS = {
    'AL': 'Alabama', 'AK': 'Alaska', 'AR': 'Arkansas', 'AZ': 'Arizona',
    'CA': 'California', 'CO': 'Colorado', 'CT': 'Connecticut', 'DE': 'Delaware',
    'DC': 'District of Columbia', 'FL': 'Florida', 'GA': 'Georgia', 'HI': 'Hawaii',
    'ID': 'Idaho', 'IL': 'Illinois', 'IN': 'Indiana', 'IA': 'Iowa',
    'KS': 'Kansas', 'KY': 'Kentucky', 'LA': 'Louisiana', 'ME': 'Maine',
    'MD': 'Maryland', 'MA': 'Massachusetts', 'MI': 'Michigan', 'MN': 'Minnesota',
    'MS': 'Mississippi', 'MO': 'Missouri', 'MT': 'Montana', 'NE': 'Nebraska',
    'NV': 'Nevada', 'NH': 'New Hampshire', 'NJ': 'New Jersey', 'NM': 'New Mexico',
    'NY': 'New York', 'NC': 'North Carolina', 'ND': 'North Dakota', 'OH': 'Ohio',
    'OK': 'Oklahoma', 'OR': 'Oregon', 'PA': 'Pennsylvania', 'RI': 'Rhode Island',
    'SC': 'South Carolina', 'SD': 'South Dakota', 'TN': 'Tennessee', 'TX': 'Texas',
    'UT': 'Utah', 'VT': 'Vermont', 'VA': 'Virginia', 'WA': 'Washington',
    'WV': 'West Virginia', 'WI': 'Wisconsin', 'WY': 'Wyoming'
}

def expand_state_abbreviations(text):
    """Expand state abbreviations to full state names - only if the entire text is an abbreviation"""
    if pd.isna(text):
        return text
    
    text_str = str(text).strip()
    
    # Check if the entire text (case-insensitive) is a state abbreviation
    # Only expand if it's exactly an abbreviation, not if abbreviation appears within a longer word
    text_upper = text_str.upper()
    if text_upper in STATE_ABBREVIATIONS:
        return STATE_ABBREVIATIONS[text_upper]
    
    # If not a direct match, return the original text unchanged
    # This preserves full state names like "Arizona", "Minnesota", etc.
    return text_str

def format_insurance_name(insurance_text):
    """Format insurance name to match expected format"""
    if pd.isna(insurance_text):
        return insurance_text
    
    insurance_str = str(insurance_text).strip()
    
    # Handle special cases first
    if insurance_str.upper() == 'NO INSURANCE':
        return 'No Insurance'
    elif insurance_str.upper() == 'PATIENT NOT FOUND':
        return 'PATIENT NOT FOUND'
    elif insurance_str.upper() == 'DUPLICATE':
        return 'DUPLICATE'
    elif re.search(r'no\s+patient\s+chart', insurance_str, re.IGNORECASE):
        return 'No Patient chart'
    
    # Extract company name before "Ph#"
    if "Ph#" in insurance_str:
        company_name = insurance_str.split("Ph#")[0].strip()
    else:
        company_name = insurance_str
    
    # Remove "Primary" and "Secondary" text
    company_name = re.sub(r'\s*\(Primary\)', '', company_name, flags=re.IGNORECASE)
    company_name = re.sub(r'\s*\(Secondary\)', '', company_name, flags=re.IGNORECASE)
    company_name = re.sub(r'\s*Primary', '', company_name, flags=re.IGNORECASE)
    company_name = re.sub(r'\s*Secondary', '', company_name, flags=re.IGNORECASE)
    
    # If already formatted as "DD [State]", preserve it (don't reformat)
    if re.match(r'^dd\s+', company_name, re.IGNORECASE):
        # Extract the state part
        state_part = re.sub(r'^dd\s+', '', company_name, flags=re.IGNORECASE).strip()
        # Remove any trailing text (like "Ph#")
        state_part = re.split(r'\s*[,\|;]\s*|\s+Ph#', state_part, flags=re.IGNORECASE)[0].strip()
        # Capitalize properly (Title Case) if not already uppercase
        if state_part and not state_part.isupper():
            state_part = state_part.title()
        return f"DD {state_part}"
    
    # Handle Delta Dental variations
    if re.search(r'delta\s+dental', company_name, re.IGNORECASE):
        # Extract state from Delta Dental - handles "Delta Dental Arizona", "Delta Dental of Arizona", etc.
        delta_match = re.search(r'delta\s+dental\s+(?:of\s+)?(.+)', company_name, re.IGNORECASE)
        if delta_match:
            state = delta_match.group(1).strip()
            # Remove any trailing text after state (like "Ph#" or other info)
            # Split on common separators and take first part
            state = re.split(r'\s*[,\|;]\s*|\s+Ph#', state, flags=re.IGNORECASE)[0].strip()
            # Expand state abbreviations (e.g., "AZ" -> "Arizona")
            state = expand_state_abbreviations(state)
            # Capitalize properly (Title Case)
            if state and not state.isupper():
                state = state.title()
            return f"DD {state}"
        else:
            return "DD"
    
    # Handle Anthem variations FIRST (before BCBS to avoid conflicts)
    if re.search(r'anthem|blue\s+cross.*anthem|anthem.*blue\s+cross', company_name, re.IGNORECASE):
        return "Anthem"
    
    # Handle BCBS variations
    elif re.search(r'bcbs|bc/bs|bc\s+of|blue\s+cross|blue\s+shield|bcbbs', company_name, re.IGNORECASE):
        # Check for full "Blue Cross Blue Shield" pattern first
        if re.search(r'blue\s+cross\s+blue\s+shield', company_name, re.IGNORECASE):
            # Extract state from "Blue Cross Blue Shield of [State]"
            bcbs_match = re.search(r'blue\s+cross\s+blue\s+shield\s+(?:of\s+)?(.+)', company_name, re.IGNORECASE)
            if bcbs_match:
                state = bcbs_match.group(1).strip()
                # Expand state abbreviations
                state = expand_state_abbreviations(state)
                return f"BCBS {state}"
            else:
                return "BCBS"
        # Handle BC/BS patterns
        elif re.search(r'bc/bs', company_name, re.IGNORECASE):
            bcbs_match = re.search(r'bc/bs\s+(?:of\s+)?(.+)', company_name, re.IGNORECASE)
            if bcbs_match:
                state = bcbs_match.group(1).strip()
                # Expand state abbreviations
                state = expand_state_abbreviations(state)
                return f"BCBS {state}"
            else:
                return "BCBS"
        # Handle BC Of patterns
        elif re.search(r'bc\s+of', company_name, re.IGNORECASE):
            bcbs_match = re.search(r'bc\s+of\s+(.+)', company_name, re.IGNORECASE)
            if bcbs_match:
                state = bcbs_match.group(1).strip()
                # Expand state abbreviations
                state = expand_state_abbreviations(state)
                return f"BCBS {state}"
            else:
                return "BCBS"
        # Handle BCBBS typo
        elif re.search(r'bcbbs', company_name, re.IGNORECASE):
            return "BCBS"
        # Handle other BCBS patterns
        else:
            bcbs_match = re.search(r'(?:bcbs|blue\s+cross|blue\s+shield)\s+(?:of\s+)?(.+)', company_name, re.IGNORECASE)
            if bcbs_match:
                state = bcbs_match.group(1).strip()
                # Expand state abbreviations
                state = expand_state_abbreviations(state)
                return f"BCBS {state}"
            else:
                return "BCBS"
    
    # Handle other specific companies
    elif re.search(r'metlife|met\s+life', company_name, re.IGNORECASE):
        return "Metlife"
    elif re.search(r'cigna', company_name, re.IGNORECASE):
        return "Cigna"
    elif re.search(r'aarp', company_name, re.IGNORECASE):
        return "AARP"
    elif re.search(r'adn\s+administrators', company_name, re.IGNORECASE):
        return "ADN Administrators"
    elif re.search(r'beam', company_name, re.IGNORECASE):
        return "Beam"
    elif re.search(r'uhc|united.*health|united.*heal|unitedhelathcare', company_name, re.IGNORECASE):
        return "UHC"
    elif re.search(r'teamcare', company_name, re.IGNORECASE):
        return "Teamcare"
    elif re.search(r'humana', company_name, re.IGNORECASE):
        return "Humana"
    elif re.search(r'aetna', company_name, re.IGNORECASE):
        return "Aetna"
    elif re.search(r'guardian', company_name, re.IGNORECASE):
        return "Guardian"
    elif re.search(r'g\s*e\s*h\s*a', company_name, re.IGNORECASE):
        return "GEHA"
    elif re.search(r'principal', company_name, re.IGNORECASE):
        return "Principal"
    elif re.search(r'ameritas', company_name, re.IGNORECASE):
        return "Ameritas"
    elif re.search(r'physicians\s+mutual', company_name, re.IGNORECASE):
        return "Physicians Mutual"
    elif re.search(r'mutual\s+of\s+omaha', company_name, re.IGNORECASE):
        return "Mutual Omaha"
    elif re.search(r'sunlife|sun\s+life', company_name, re.IGNORECASE):
        return "Sunlife"
    elif re.search(r'liberty(?:\s+dental)?', company_name, re.IGNORECASE):
        return "Liberty Dental Plan"
    elif re.search(r'careington', company_name, re.IGNORECASE):
        return "Careington Benefit Solutions"
    elif re.search(r'automated\s+benefit', company_name, re.IGNORECASE):
        return "Automated Benefit Services Inc"
    elif re.search(r'network\s+health', company_name, re.IGNORECASE):
        # Check if it has "Wisconsin" in the name
        if re.search(r'wisconsin', company_name, re.IGNORECASE):
            return "Network Health Wisconsin"
        else:
            return "Network Health Go"
    elif re.search(r'regence', company_name, re.IGNORECASE):
        return "REGENCE BCBS"
    elif re.search(r'united\s+concordia', company_name, re.IGNORECASE):
        return "United Concordia"
    elif re.search(r'medical\s+mutual', company_name, re.IGNORECASE):
        return "Medical Mutual"
    elif re.search(r'blue\s+care\s+dental', company_name, re.IGNORECASE):
        return "Blue Care Dental"
    elif re.search(r'dominion\s+dental', company_name, re.IGNORECASE):
        return "Dominion Dental"
    elif re.search(r'carefirst', company_name, re.IGNORECASE):
        return "CareFirst BCBS"
    elif re.search(r'health\s*partners', company_name, re.IGNORECASE):
        # Check if it has "of [State]" pattern
        if re.search(r'health\s*partners\s+of\s+(.+)', company_name, re.IGNORECASE):
            state_match = re.search(r'health\s*partners\s+of\s+(.+)', company_name, re.IGNORECASE)
            state = state_match.group(1).strip()
            return f"Health Partners {state}"
        else:
            return "Health Partners"
    elif re.search(r'keenan', company_name, re.IGNORECASE):
        return "Keenan"
    elif re.search(r'wilson\s+mcshane', company_name, re.IGNORECASE):
        return "Wilson McShane- Delta Dental"
    elif re.search(r'standard\s+(?:life\s+)?insurance', company_name, re.IGNORECASE):
        return "Standard Life Insurance"
    elif re.search(r'plan\s+for\s+health', company_name, re.IGNORECASE):
        return "Plan for Health"
    elif re.search(r'kansas\s+city', company_name, re.IGNORECASE):
        return "Kansas City"
    elif re.search(r'the\s+guardian', company_name, re.IGNORECASE):
        return "The Guardian"
    elif re.search(r'community\s+dental', company_name, re.IGNORECASE):
        return "Community Dental Associates"
    elif re.search(r'northeast\s+delta\s+dental', company_name, re.IGNORECASE):
        return "Northeast Delta Dental"
    elif re.search(r'say\s+cheese\s+dental', company_name, re.IGNORECASE):
        return "Say Cheese Dental Network"
    elif re.search(r'dentaquest', company_name, re.IGNORECASE):
        return "Dentaquest"
    elif re.search(r'umr', company_name, re.IGNORECASE):
        return "UMR"
    elif re.search(r'mhbp', company_name, re.IGNORECASE):
        return "MHBP"
    elif re.search(r'united\s+states\s+army', company_name, re.IGNORECASE):
        return "United States Army"
    elif re.search(r'conversion\s+default', company_name, re.IGNORECASE):
        return "CONVERSION DEFAULT - Do NOT Delete! Change Pt Ins!"
    elif re.search(r'equitable', company_name, re.IGNORECASE):
        return "Equitable"
    elif re.search(r'manhattan\s+life', company_name, re.IGNORECASE):
        return "Manhattan Life"
    elif re.search(r'ucci', company_name, re.IGNORECASE):
        return "UCCI"
    elif re.search(r'ccpoa|cc\s*poa|c\s+c\s+p\s+o\s+a', company_name, re.IGNORECASE):
        return "CCPOA"
    elif re.search(r'dd\s+of|dd\s+[a-z]{2}|delta\s+dental|dental\s+dental|denta\s+dental|dleta\s+dental|dektal?\s+dental', company_name, re.IGNORECASE):
        # Extract state from various Delta Dental patterns
        # Handle DD OF [State] pattern
        if re.search(r'dd\s+of\s+([a-z]{2})', company_name, re.IGNORECASE):
            state_match = re.search(r'dd\s+of\s+([a-z]{2})', company_name, re.IGNORECASE)
            state = state_match.group(1).upper()
            state = expand_state_abbreviations(state)
            return f'DD {state}'
        # Handle DD [State] pattern - only match if it's exactly 2 letters (abbreviation) followed by word boundary
        elif re.search(r'dd\s+([a-z]{2})\b', company_name, re.IGNORECASE):
            state_match = re.search(r'dd\s+([a-z]{2})\b', company_name, re.IGNORECASE)
            state = state_match.group(1).upper()
            state = expand_state_abbreviations(state)
            return f'DD {state}'
        # Handle DD [Full State Name] pattern - for cases like "DD Pennsylvania"
        elif re.search(r'^dd\s+([a-z]{3,})', company_name, re.IGNORECASE):
            state_match = re.search(r'^dd\s+([a-z]{3,})', company_name, re.IGNORECASE)
            state = state_match.group(1).strip()
            # Remove any trailing text after state (like "Ph#" or other info)
            state = re.split(r'\s*[,\|;]\s*|\s+Ph#', state, flags=re.IGNORECASE)[0].strip()
            # Don't expand abbreviations - preserve full state name
            # Capitalize properly (Title Case)
            if state and not state.isupper():
                state = state.title()
            return f'DD {state}'
        # Handle Delta Dental of [State] pattern
        elif re.search(r'delta\s+dental\s+of\s+(.+)', company_name, re.IGNORECASE):
            state_match = re.search(r'delta\s+dental\s+of\s+(.+)', company_name, re.IGNORECASE)
            state = state_match.group(1).strip()
            state = expand_state_abbreviations(state)
            return f'DD {state}'
        # Handle Dental Dental Of [State] pattern
        elif re.search(r'dental\s+dental\s+of\s+(.+)', company_name, re.IGNORECASE):
            state_match = re.search(r'dental\s+dental\s+of\s+(.+)', company_name, re.IGNORECASE)
            state = state_match.group(1).strip()
            state = expand_state_abbreviations(state)
            return f'DD {state}'
        # Handle Denta Dental Of [State] pattern
        elif re.search(r'denta\s+dental\s+of\s+(.+)', company_name, re.IGNORECASE):
            state_match = re.search(r'denta\s+dental\s+of\s+(.+)', company_name, re.IGNORECASE)
            state = state_match.group(1).strip()
            state = expand_state_abbreviations(state)
            return f'DD {state}'
        # Handle Dleta Dental of [State] pattern
        elif re.search(r'dleta\s+dental\s+of\s+(.+)', company_name, re.IGNORECASE):
            state_match = re.search(r'dleta\s+dental\s+of\s+(.+)', company_name, re.IGNORECASE)
            state = state_match.group(1).strip()
            state = expand_state_abbreviations(state)
            return f'DD {state}'
        # Handle Dekta Dental of [State] pattern
        elif re.search(r'dektal?\s+dental\s+of\s+(.+)', company_name, re.IGNORECASE):
            state_match = re.search(r'dektal?\s+dental\s+of\s+(.+)', company_name, re.IGNORECASE)
            state = state_match.group(1).strip()
            state = expand_state_abbreviations(state)
            return f'DD {state}'
        # Handle Dental of [State] pattern
        elif re.search(r'dental\s+of\s+(.+)', company_name, re.IGNORECASE):
            state_match = re.search(r'dental\s+of\s+(.+)', company_name, re.IGNORECASE)
            state = state_match.group(1).strip()
            state = expand_state_abbreviations(state)
            return f'DD {state}'
        # Handle Dental Network of America
        elif re.search(r'dental\s+network\s+of\s+america', company_name, re.IGNORECASE):
            return 'DD Network of America'
        # Default DD
        else:
            return 'DD'
    
    # If no specific pattern matches, return the cleaned company name
    return company_name.strip()

def compare_patient_names(raw_df, previous_df):
    """Compare Patient ID from Appointment report with PATID from Smart Assist and add insurance columns"""
    try:
        # Debug: Show available columns
        raw_columns = list(raw_df.columns)
        previous_columns = list(previous_df.columns)
        
        # Function to find specific column names
        def find_column(columns, target_names):
            """Find column matching any of the target names (case-insensitive)"""
            for col in columns:
                col_lower = col.lower().strip()
                for target in target_names:
                    if col_lower == target.lower().strip():
                        return col
            return None
        
        # Find Patient ID in raw file (Appointment report)
        raw_patient_col = find_column(raw_columns, ['Patient ID', 'PatientID', 'patient id'])
        if not raw_patient_col:
            # Try flexible search as fallback
            for col in raw_columns:
                if 'patient' in col.lower() and 'id' in col.lower():
                    raw_patient_col = col
                    break
        
        # Find PATID in previous file (Smart Assist)
        previous_patient_col = find_column(previous_columns, ['PATID', 'PatID', 'patid', 'PAT ID'])
        if not previous_patient_col:
            return f"❌ Error: 'PATID' column not found in Smart Assist file.\nAvailable columns: {previous_columns}", None
        
        # Find insurance columns in Smart Assist file, or create them if missing
        primary_ins_col = find_column(previous_columns, ['Dental Primary Ins Carr', 'DentalPrimaryInsCarr', 'Dental Primary Insurance Carrier'])
        secondary_ins_col = find_column(previous_columns, ['Dental Secondary Ins Carr', 'DentalSecondaryInsCarr', 'Dental Secondary Insurance Carrier'])
        
        # Create a working copy of the Smart Assist DataFrame
        smart_assist_df = previous_df.copy()
        
        # Track if columns were added
        columns_added = []
        
        # Add missing insurance columns if they don't exist
        if not primary_ins_col:
            primary_ins_col = 'Dental Primary Ins Carr'
            smart_assist_df[primary_ins_col] = ''
            columns_added.append(primary_ins_col)
        
        if not secondary_ins_col:
            secondary_ins_col = 'Dental Secondary Ins Carr'
            smart_assist_df[secondary_ins_col] = ''
            columns_added.append(secondary_ins_col)
        
        # Check if Patient ID column exists in raw file
        if not raw_patient_col:
            return f"❌ Error: 'Patient ID' column not found in Appointment report file.\nAvailable columns: {raw_columns}", None
        
        # Find insurance columns in Appointment Report file (more flexible search)
        appointment_primary_col = None
        appointment_secondary_col = None
        
        # Try exact matches first
        appointment_primary_col = find_column(raw_columns, ['Dental Primary Ins Carr', 'DentalPrimaryInsCarr', 'Dental Primary Insurance Carrier'])
        appointment_secondary_col = find_column(raw_columns, ['Dental Secondary Ins Carr', 'DentalSecondaryInsCarr', 'Dental Secondary Insurance Carrier'])
        
        # If not found, try partial matches
        if not appointment_primary_col:
            for col in raw_columns:
                col_lower = col.lower()
                if ('primary' in col_lower and 'ins' in col_lower) or ('primary' in col_lower and 'insurance' in col_lower):
                    appointment_primary_col = col
                    break
        
        if not appointment_secondary_col:
            for col in raw_columns:
                col_lower = col.lower()
                if ('secondary' in col_lower and 'ins' in col_lower) or ('secondary' in col_lower and 'insurance' in col_lower):
                    appointment_secondary_col = col
                    break
        
        # Create a mapping from Patient ID to insurance data in Appointment Report
        appointment_insurance_map = {}
        for idx, row in raw_df.iterrows():
            patient_id_val = row[raw_patient_col]
            if pd.notna(patient_id_val):
                # Convert to string and strip for consistent matching
                patient_id = str(patient_id_val).strip()
                if patient_id:
                    primary_value = ''
                    secondary_value = ''
                    
                    if appointment_primary_col:
                        primary_val = row[appointment_primary_col]
                        if pd.notna(primary_val):
                            primary_value = str(primary_val).strip()
                    
                    if appointment_secondary_col:
                        secondary_val = row[appointment_secondary_col]
                        if pd.notna(secondary_val):
                            secondary_value = str(secondary_val).strip()
                    
                    appointment_insurance_map[patient_id] = {
                        'primary': primary_value,
                        'secondary': secondary_value
                    }
        
        # Use Smart Assist file as the result file (base)
        result_df = smart_assist_df.copy()
        
        # Find the position of the PATID column in Smart Assist
        patid_col_index = result_df.columns.get_loc(previous_patient_col)
        
        # Check if insurance columns already exist, if not create them
        primary_col_name = 'Dental Primary Ins Carr'
        secondary_col_name = 'Dental Secondary Ins Carr'
        
        if primary_col_name not in result_df.columns:
            result_df.insert(patid_col_index + 1, primary_col_name, '')
        else:
            # Column already exists, initialize empty values if needed
            result_df[primary_col_name] = result_df[primary_col_name].fillna('')
        
        if secondary_col_name not in result_df.columns:
            # Insert after primary column (or after PATID if primary doesn't exist)
            if primary_col_name in result_df.columns:
                # Find the position of primary column and insert after it
                primary_col_index = result_df.columns.get_loc(primary_col_name)
                result_df.insert(primary_col_index + 1, secondary_col_name, '')
            else:
                # Primary column doesn't exist, insert after PATID
                result_df.insert(patid_col_index + 1, secondary_col_name, '')
        else:
            # Column already exists, initialize empty values if needed
            result_df[secondary_col_name] = result_df[secondary_col_name].fillna('')
        
        # Compare and populate insurance columns in Smart Assist file from Appointment Report
        matched_count = 0
        for idx, row in result_df.iterrows():
            patid_val = row[previous_patient_col]
            if pd.notna(patid_val):
                # Convert to string and strip for consistent matching
                patid = str(patid_val).strip()
                if patid and patid in appointment_insurance_map:
                    # Match found - Patient ID from Appointment Report matches PATID in Smart Assist
                    # Copy insurance data FROM Appointment Report TO Smart Assist file
                    # Format the insurance names before adding
                    insurance_data = appointment_insurance_map[patid]
                    result_df.at[idx, primary_col_name] = format_insurance_name(insurance_data['primary'])
                    result_df.at[idx, secondary_col_name] = format_insurance_name(insurance_data['secondary'])
                    matched_count += 1
        
        # Count statistics
        total_patients = len(result_df)
        not_matched_patients = total_patients - matched_count
        
        # Generate result message
        columns_info = ""
        if columns_added:
            columns_info = f"\n⚠️ Note: The following columns were added to Smart Assist file (with empty values): {', '.join(columns_added)}"
        
        result_message = f"""✅ Comparison completed successfully!

📊 Statistics:
- Total records in Smart Assist file: {total_patients}
- Records matched with Appointment report: {matched_count}
- Records not matched: {not_matched_patients}
- Match rate: {(matched_count/total_patients*100):.1f}%

📋 Columns used:
- Appointment report (Raw file): "{raw_patient_col}"
  - Primary Insurance: {appointment_primary_col if appointment_primary_col else "❌ Not found - will use empty values"}
  - Secondary Insurance: {appointment_secondary_col if appointment_secondary_col else "❌ Not found - will use empty values"}
- Smart Assist file (Previous file): "{previous_patient_col}"
  - Insurance columns: "{primary_ins_col}", "{secondary_ins_col}"{columns_info}
  
📊 Matching details:
- Total Patient IDs in Appointment Report: {len(appointment_insurance_map)}
- Records in Smart Assist: {total_patients}
- Successful matches: {matched_count}

📋 Result file:
- Based on Smart Assist file
- "Dental Primary Ins Carr" - populated from Appointment Report for matched records
- "Dental Secondary Ins Carr" - populated from Appointment Report for matched records
- Empty cells for records not found in Appointment report

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
        # Get filename without saving to disk
        filename = secure_filename(file.filename)
        
        # Read Excel file directly from memory (no disk storage)
        file.seek(0)  # Reset file pointer to beginning
        raw_data = pd.read_excel(file, sheet_name=None, engine='openpyxl')
        raw_filename = filename
        
        comparison_result = f"✅ Appointment Report uploaded successfully! Loaded {len(raw_data)} sheets: {', '.join(list(raw_data.keys()))}"
        return redirect('/comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error uploading Appointment Report: {str(e)}"
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
        # Get filename without saving to disk
        filename = secure_filename(file.filename)
        
        # Read Excel file directly from memory (no disk storage)
        file.seek(0)  # Reset file pointer to beginning
        previous_data = pd.read_excel(file, sheet_name=None, engine='openpyxl')
        previous_filename = filename
        
        comparison_result = f"✅ Smart Assist file uploaded successfully! Loaded {len(previous_data)} sheets: {', '.join(list(previous_data.keys()))}"
        return redirect('/comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error uploading Smart Assist file: {str(e)}"
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
            # Update the previous_data with the result (result is based on Smart Assist file)
            previous_data[previous_sheet] = result_df
        else:
            comparison_result = result_message
        
        return redirect('/comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error comparing files: {str(e)}"
        return redirect('/comparison')

@app.route('/download_result', methods=['POST'])
def download_result():
    global previous_data, previous_filename
    
    if not previous_data:
        return jsonify({'error': 'No data to download'}), 400
    
    filename = request.form.get('filename', '').strip()
    if not filename:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"smart_assist_result_{timestamp}.xlsx"
    
    try:
        # Create a temporary file
        import tempfile
        temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
        
        try:
            with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                for sheet_name, df in previous_data.items():
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
