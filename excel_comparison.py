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

# Global variables for conversion report
conversion_data = None
conversion_filename = None
conversion_result = None

# Global variables for insurance name formatting
insurance_formatting_data = None
insurance_formatting_filename = None
insurance_formatting_result = None
insurance_formatting_output = ""

# Global variables for remarks update
remarks_appointments_data = None
remarks_excel_data = None
remarks_appointments_filename = None
remarks_remarks_filename = None
remarks_result = None
remarks_updated_count = 0

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
        
        /* Processing Modal Overlay */
        .processing-modal-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.7);
            backdrop-filter: blur(5px);
            z-index: 9999;
            justify-content: center;
            align-items: center;
            animation: fadeIn 0.3s ease;
        }
        
        .processing-modal-overlay.show {
            display: flex;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
        }
        
        .processing-modal {
            background: white;
            border-radius: 20px;
            padding: 40px;
            max-width: 400px;
            width: 90%;
            text-align: center;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            animation: slideUp 0.3s ease;
        }
        
        @keyframes slideUp {
            from {
                transform: translateY(30px);
                opacity: 0;
            }
            to {
                transform: translateY(0);
                opacity: 1;
            }
        }
        
        .processing-modal .modal-spinner {
            border: 5px solid #f3f3f3;
            border-top: 5px solid #667eea;
            border-radius: 50%;
            width: 60px;
            height: 60px;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }
        
        .processing-modal h3 {
            color: #333;
            margin-bottom: 15px;
            font-size: 1.5em;
        }
        
        .processing-modal p {
            color: #666;
            font-size: 1.1em;
            margin: 0;
        }
        
        .processing-modal .processing-dots {
            display: inline-block;
            animation: dots 1.5s steps(4, end) infinite;
        }
        
        .processing-modal .processing-dots::after {
            content: '...';
            animation: dots 1.5s steps(4, end) infinite;
        }
        
        @keyframes dots {
            0%, 20% { content: '.'; }
            40% { content: '..'; }
            60%, 100% { content: '...'; }
        }
        
        /* Tab Styles */
        .tabs {
            display: flex;
            border-bottom: 2px solid #e0e0e0;
            margin-bottom: 20px;
            background: transparent;
            padding: 5px 0;
        }
        .tab {
            padding: 15px 30px;
            cursor: pointer;
            background: #f8f9fa;
            border: none;
            border-bottom: 3px solid transparent;
            font-size: 16px;
            font-weight: 600;
            color: #495057;
            transition: all 0.3s;
            border-radius: 8px 8px 0 0;
            margin-right: 5px;
        }
        .tab:hover {
            background: #e9ecef;
            color: #212529;
        }
        .tab.active {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #ffffff;
            border-bottom-color: #667eea;
            box-shadow: 0 -2px 8px rgba(102, 126, 234, 0.3);
            border-bottom-width: 4px;
        }
        .tab-content {
            display: none;
        }
        .tab-content.active {
            display: block;
        }
        
        /* Toast Notification */
        .toast {
            position: fixed;
            top: 20px;
            right: 20px;
            background: #4caf50;
            color: white;
            padding: 16px 24px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 10000;
            display: none;
            animation: slideIn 0.3s ease;
            max-width: 400px;
        }
        .toast.error {
            background: #f44336;
        }
        .toast.show {
            display: block;
        }
        @keyframes slideIn {
            from {
                transform: translateX(400px);
                opacity: 0;
            }
            to {
                transform: translateX(0);
                opacity: 1;
            }
        }
        .toast-close {
            float: right;
            font-weight: bold;
            cursor: pointer;
            margin-left: 15px;
        }
    </style>
</head>
<body>
    <!-- Toast Notification -->
    <div id="toast" class="toast">
        <span id="toast-message"></span>
        <span class="toast-close" onclick="hideToast()">&times;</span>
    </div>

    <!-- Processing Modal Overlay -->
    <div id="processing-modal-overlay" class="processing-modal-overlay">
        <div class="processing-modal">
            <div class="modal-spinner"></div>
            <h3 id="processing-title">Processing</h3>
            <p id="processing-message">Please wait while we process your request<span class="processing-dots">...</span></p>
        </div>
    </div>

    <div class="container">
        <div class="header">
            <h1>📊 Excel Automation Tools</h1>
            <p>Multiple Excel processing tools in one place</p>
        </div>

        <div class="content">
            <!-- Tab Navigation -->
            <div class="tabs">
                <button class="tab {% if active_tab == 'comparison' %}active{% endif %}" id="comparison-tab-btn" onclick="switchTab('comparison')">🔄 Comparison Tool</button>
                <button class="tab {% if active_tab == 'conversion' %}active{% endif %}" id="conversion-tab-btn" onclick="switchTab('conversion')">📋 Conversion Report Formatting</button>
                <button class="tab {% if active_tab == 'insurance' %}active{% endif %}" id="insurance-tab-btn" onclick="switchTab('insurance')">🦷 Insurance Name Formatting</button>
                <button class="tab {% if active_tab == 'remarks' %}active{% endif %}" id="remarks-tab-btn" onclick="switchTab('remarks')">📝 Update Remarks</button>
            </div>

            <!-- Tab 1: Comparison Tool -->
            <div id="comparison-tab" class="tab-content {% if active_tab == 'comparison' %}active{% endif %}">
                <div class="section">
                    <h3>🔄 Comparison Tool</h3>
                    <p>Compare Patient ID from Appointment Report with PATID from Smart Assist file. When matched, insurance columns will be added to the Smart Assist file.</p>
                </div>

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
                <h3>🔄 Reset Comparison Tool</h3>
                <p>Clear all uploaded files and reset the comparison tool to start fresh.</p>
                <form action="/reset_comparison" method="post" onsubmit="return confirm('Are you sure you want to reset the comparison tool? This will clear all uploaded files and data.')">
                    <button type="submit" class="reset-btn">🗑️ Reset Comparison Tool</button>
                </form>
            </div>
            </div>

            <!-- Tab 2: Conversion Report Formatting -->
            <div id="conversion-tab" class="tab-content {% if active_tab == 'conversion' %}active{% endif %}">
                <div class="section">
                    <h3>📋 Conversion Report Formatting</h3>
                    <p>Upload an Excel file to process and format conversion reports. The file must contain an "Insurance Note" column.</p>
                </div>

                <!-- File Upload Section -->
                <div class="section">
                    <h3>📁 Upload Excel File</h3>
                    <form action="/upload_conversion" method="post" enctype="multipart/form-data" id="conversion-form">
                        <div class="form-group">
                            <label for="conversion_file">Select Conversion Report Excel File:</label>
                            <input type="file" id="conversion_file" name="file" accept=".xlsx,.xls" required>
                        </div>
                        <button type="submit" id="conversion-btn">📤 Upload & Validate File</button>
                    </form>
                    <div class="loading" id="conversion-loading">
                        <div class="spinner"></div>
                        <p>Processing conversion report...</p>
                    </div>
                </div>

                <!-- Status Messages -->
                {% if conversion_result %}
                <div class="section">
                    <h3>📢 Processing Status</h3>
                    <div class="status-message">
                        {{ conversion_result | safe }}
                    </div>
                </div>
                {% endif %}

                <!-- File Status -->
                {% if conversion_filename %}
                <div class="section">
                    <h3>📊 File Status</h3>
                    <div class="file-status">
                        <div class="status-success">
                            ✅ Conversion Report: {{ conversion_filename }}<br>
                            📋 Sheets: {{ conversion_data.keys() | list | length if conversion_data else 0 }}
                        </div>
                    </div>
                </div>
                {% endif %}

                <!-- Download Section -->
                {% if conversion_data and conversion_result and 'processing completed successfully' in conversion_result.lower() %}
                <div class="section">
                    <h3>💾 Download Processed File</h3>
                    <form action="/download_conversion" method="post">
                        <div class="form-group">
                            <label for="conversion_output_filename">Output filename (optional):</label>
                            <input type="text" id="conversion_output_filename" name="filename" 
                                   placeholder="conversion_report_formatted.xlsx" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                        </div>
                        <button type="submit">💾 Download Processed File</button>
                    </form>
                </div>
                {% endif %}

                <!-- Reset Section -->
                <div class="section">
                    <h3>🔄 Reset Conversion Tool</h3>
                    <p>Clear all uploaded files and reset the conversion tool to start fresh.</p>
                    <form action="/reset_conversion" method="post" onsubmit="return confirm('Are you sure you want to reset the conversion tool? This will clear all uploaded files and data.')">
                        <button type="submit" class="reset-btn">🗑️ Reset Conversion Tool</button>
                    </form>
                </div>
            </div>

            <!-- Tab 3: Insurance Name Formatting -->
            <div id="insurance-tab" class="tab-content {% if active_tab == 'insurance' %}active{% endif %}">
                <div class="section">
                    <h3>🦷 Insurance Name Formatting</h3>
                    <p>Upload an Excel file to automatically format "Dental Primary Ins Carr" column in all sheets. A new formatted column will be created.</p>
                </div>

                <!-- File Upload Section -->
                <div class="section">
                    <h3>📁 Upload Excel File</h3>
                    <form action="/upload_insurance_formatting" method="post" enctype="multipart/form-data" id="insurance-form">
                        <div class="form-group">
                            <label for="insurance_file">Select Excel File (.xlsx, .xls):</label>
                            <input type="file" id="insurance_file" name="file" accept=".xlsx,.xls" required>
                        </div>
                        <button type="submit" id="insurance-btn">📤 Upload & Process File</button>
                    </form>
                    <div class="loading" id="insurance-loading">
                        <div class="spinner"></div>
                        <p>Processing file and formatting columns...</p>
                    </div>
                </div>

                <!-- Status Messages -->
                {% if insurance_formatting_result %}
                <div class="section">
                    <h3>📢 Processing Status</h3>
                    <div class="status-message">
                        {{ insurance_formatting_result | safe }}
                    </div>
                </div>
                {% endif %}

                <!-- File Status -->
                {% if insurance_formatting_filename %}
                <div class="section">
                    <h3>📊 File Status</h3>
                    <div class="file-status">
                        <div class="status-success">
                            ✅ File loaded: {{ insurance_formatting_filename }}<br>
                            📋 Sheets processed: {{ insurance_formatting_data.keys() | list | length if insurance_formatting_data else 0 }}
                        </div>
                    </div>
                </div>
                {% endif %}

                <!-- Processing Output -->
                {% if insurance_formatting_output %}
                <div class="section">
                    <h3>📝 Processing Output</h3>
                    <div class="output" style="background: #1e1e1e; color: #f8f8f2; padding: 20px; border-radius: 8px; white-space: pre-wrap; font-family: 'Courier New', monospace; max-height: 500px; overflow-y: auto; border: 1px solid #333; font-size: 14px;">
                        {{ insurance_formatting_output }}
                    </div>
                </div>
                {% endif %}

                <!-- Download Section -->
                {% if insurance_formatting_data and insurance_formatting_result and 'processing complete' in insurance_formatting_result.lower() %}
                <div class="section">
                    <h3>💾 Download Formatted File</h3>
                    <form action="/download_insurance_formatting" method="post">
                        <div class="form-group">
                            <label for="insurance_output_filename">Output filename (optional):</label>
                            <input type="text" id="insurance_output_filename" name="filename" 
                                   placeholder="formatted_insurance_names.xlsx" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                        </div>
                        <button type="submit">💾 Download Formatted File</button>
                    </form>
                </div>
                {% endif %}

                <!-- Reset Section -->
                <div class="section">
                    <h3>🔄 Reset Insurance Formatting Tool</h3>
                    <p>Clear all uploaded files and reset the insurance formatting tool to start fresh.</p>
                    <form action="/reset_insurance_formatting" method="post" onsubmit="return confirm('Are you sure you want to reset the insurance formatting tool? This will clear all uploaded files and data.')">
                        <button type="submit" class="reset-btn">🗑️ Reset Insurance Formatting Tool</button>
                    </form>
                </div>
            </div>

            <!-- Tab 4: Update Remarks -->
            <div id="remarks-tab" class="tab-content {% if active_tab == 'remarks' %}active{% endif %}">
                <div class="section">
                    <h3>📝 Update Remarks</h3>
                    <p>Upload two Excel files: Appointments file (with Pat ID) and Remarks file (with Patient ID, Remark, and Agent Name). The tool will match Patient IDs and update appointments with remarks.</p>
                </div>

                <!-- File Upload Section -->
                <div class="section">
                    <h3>📁 Upload Excel Files</h3>
                    <form action="/upload_remarks" method="post" enctype="multipart/form-data" id="remarks-form">
                        <div class="two-column">
                            <div>
                                <h4>Appointments Excel File</h4>
                                <div class="form-group">
                                    <label for="appointments_file">Select Appointments Excel File (must have "Pat ID" column):</label>
                                    <input type="file" id="appointments_file" name="appointments_file" accept=".xlsx,.xls" required>
                                </div>
                                {% if remarks_appointments_filename %}
                                <div class="file-status">
                                    <div class="status-success">
                                        ✅ {{ remarks_appointments_filename }}
                                    </div>
                                </div>
                                {% endif %}
                            </div>
                            
                            <div>
                                <h4>Remarks Excel File</h4>
                                <div class="form-group">
                                    <label for="remarks_file">Select Remarks Excel File (must have "Patient ID", "Remark", and "Agent Name" columns):</label>
                                    <input type="file" id="remarks_file" name="remarks_file" accept=".xlsx,.xls" required>
                                </div>
                                {% if remarks_remarks_filename %}
                                <div class="file-status">
                                    <div class="status-success">
                                        ✅ {{ remarks_remarks_filename }}
                                    </div>
                                </div>
                                {% endif %}
                            </div>
                        </div>
                        <button type="submit" id="remarks-btn">📤 Upload & Process Files</button>
                    </form>
                    <div class="loading" id="remarks-loading">
                        <div class="spinner"></div>
                        <p>Processing files and matching Patient IDs...</p>
                    </div>
                </div>

                <!-- Status Messages -->
                {% if remarks_result %}
                <div class="section">
                    <h3>📢 Processing Status</h3>
                    <div class="status-message">
                        {{ remarks_result | safe }}
                    </div>
                </div>
                {% endif %}

                <!-- File Status -->
                {% if remarks_appointments_data %}
                <div class="section">
                    <h3>📊 Processing Results</h3>
                    <div class="file-status">
                        <div class="status-success">
                            ✅ Total appointments processed: {{ remarks_appointments_data | length }}<br>
                            ✅ Appointments updated with remarks: {{ remarks_updated_count }}<br>
                            {% if remarks_appointments_filename %}
                            📄 Appointments file: {{ remarks_appointments_filename }}<br>
                            {% endif %}
                            {% if remarks_remarks_filename %}
                            📄 Remarks file: {{ remarks_remarks_filename }}
                            {% endif %}
                        </div>
                    </div>
                </div>
                {% endif %}

                <!-- Download Section -->
                {% if remarks_appointments_data and remarks_result and 'successfully' in remarks_result.lower() %}
                <div class="section">
                    <h3>💾 Download Updated File</h3>
                    <form action="/download_remarks" method="post">
                        <div class="form-group">
                            <label for="remarks_output_filename">Output filename (optional):</label>
                            <input type="text" id="remarks_output_filename" name="filename" 
                                   placeholder="appointments_with_remarks.xlsx" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                        </div>
                        <button type="submit">💾 Download Updated File</button>
                    </form>
                </div>
                {% endif %}

                <!-- Reset Section -->
                <div class="section">
                    <h3>🔄 Reset Remarks Tool</h3>
                    <p>Clear all uploaded files and reset the remarks tool to start fresh.</p>
                    <form action="/reset_remarks" method="post" onsubmit="return confirm('Are you sure you want to reset the remarks tool? This will clear all uploaded files and data.')">
                        <button type="submit" class="reset-btn">🗑️ Reset Remarks Tool</button>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <script>
        // Toast notification functions
        function showToast(message, isError = false) {
            const toast = document.getElementById('toast');
            const toastMessage = document.getElementById('toast-message');
            toastMessage.textContent = message;
            toast.className = 'toast' + (isError ? ' error' : '') + ' show';
            setTimeout(hideToast, 5000);
        }
        
        function hideToast() {
            const toast = document.getElementById('toast');
            toast.classList.remove('show');
        }
        
        // Processing modal functions
        function showProcessingModal(title = 'Processing', message = 'Please wait while we process your request') {
            const overlay = document.getElementById('processing-modal-overlay');
            const titleEl = document.getElementById('processing-title');
            const messageEl = document.getElementById('processing-message');
            
            titleEl.textContent = title;
            messageEl.innerHTML = message + '<span class="processing-dots">...</span>';
            overlay.classList.add('show');
            
            // Prevent body scroll when modal is open
            document.body.style.overflow = 'hidden';
        }
        
        function hideProcessingModal() {
            const overlay = document.getElementById('processing-modal-overlay');
            overlay.classList.remove('show');
            
            // Restore body scroll
            document.body.style.overflow = '';
        }
        
        // Tab switching function
        function switchTab(tabName, skipUrlUpdate = false) {
            if (!tabName) return; // Safety check
            
            // Hide all tab contents
            document.querySelectorAll('.tab-content').forEach(content => {
                if (content && content.classList) {
                    content.classList.remove('active');
                }
            });
            
            // Remove active class from all tabs
            document.querySelectorAll('.tab').forEach(tab => {
                if (tab && tab.classList) {
                    tab.classList.remove('active');
                }
            });
            
            // Show selected tab content
            const tabContent = document.getElementById(tabName + '-tab');
            if (tabContent && tabContent.classList) {
                tabContent.classList.add('active');
            }
            
            // Add active class to clicked tab
            if (typeof event !== 'undefined' && event && event.target && event.target.classList) {
                event.target.classList.add('active');
            } else {
                // If called programmatically, find the right tab button by ID
                const tabButton = document.getElementById(tabName + '-tab-btn');
                if (tabButton && tabButton.classList) {
                    tabButton.classList.add('active');
                } else {
                    // Fallback: find by text content
                    const tabs = document.querySelectorAll('.tab');
                    tabs.forEach(tab => {
                        if (tab && tab.classList && tab.textContent) {
                            const tabText = tab.textContent.toLowerCase();
                            if ((tabName === 'comparison' && tabText.includes('comparison')) ||
                                (tabName === 'conversion' && tabText.includes('conversion')) ||
                                (tabName === 'insurance' && tabText.includes('insurance')) ||
                                (tabName === 'remarks' && tabText.includes('remarks'))) {
                                tab.classList.add('active');
                            }
                        }
                    });
                }
            }
            
            // Update URL without reloading (only if not called from page load)
            if (!skipUrlUpdate) {
                const newUrl = window.location.pathname + '?tab=' + tabName;
                window.history.pushState({path: newUrl}, '', newUrl);
            }
        }
        // Form submission with processing modal
        const rawForm = document.getElementById('raw-form');
        if (rawForm) {
            rawForm.addEventListener('submit', function() {
                showProcessingModal('Uploading Appointment Report', 'Processing your Appointment Report file');
                const btn = document.getElementById('raw-btn');
                if (btn) btn.disabled = true;
            });
        }

        const previousForm = document.getElementById('previous-form');
        if (previousForm) {
            previousForm.addEventListener('submit', function() {
                showProcessingModal('Uploading Smart Assist File', 'Processing your Smart Assist file');
                const btn = document.getElementById('previous-btn');
                if (btn) btn.disabled = true;
            });
        }

        const compareForm = document.getElementById('compare-form');
        if (compareForm) {
            compareForm.addEventListener('submit', function() {
                showProcessingModal('Comparing Files', 'Matching Patient IDs and updating insurance columns');
                const btn = document.getElementById('compare-btn');
                if (btn) btn.disabled = true;
            });
        }

        // Conversion form submission
        const conversionForm = document.getElementById('conversion-form');
        if (conversionForm) {
            conversionForm.addEventListener('submit', function() {
                showProcessingModal('Processing Conversion Report', 'Validating and formatting conversion report');
                document.getElementById('conversion-btn').disabled = true;
            });
        }

        // Insurance formatting form submission
        const insuranceForm = document.getElementById('insurance-form');
        if (insuranceForm) {
            insuranceForm.addEventListener('submit', function() {
                showProcessingModal('Formatting Insurance Names', 'Processing file and formatting insurance columns');
                document.getElementById('insurance-btn').disabled = true;
            });
        }

        // Remarks form submission
        const remarksForm = document.getElementById('remarks-form');
        if (remarksForm) {
            remarksForm.addEventListener('submit', function() {
                showProcessingModal('Processing Files', 'Matching Patient IDs and updating appointments with remarks');
                document.getElementById('remarks-btn').disabled = true;
            });
        }

        // Reset forms - show modal when form is submitted (HTML confirm already handled)
        const resetForms = document.querySelectorAll('form[action*="reset"]');
        resetForms.forEach(form => {
            form.addEventListener('submit', function() {
                // Show modal - HTML onsubmit confirm already handled the confirmation
                showProcessingModal('Resetting', 'Clearing all data and files');
            });
        });

        // Auto-scroll results to bottom
        function scrollResults() {
            const results = document.getElementById('results');
            results.scrollTop = results.scrollHeight;
        }

        // Scroll results on page load
        window.onload = function() {
            scrollResults();
            
            // Hide processing modal on page load (in case it was left open)
            hideProcessingModal();
            
            // Check URL parameter for active tab
            const urlParams = new URLSearchParams(window.location.search);
            const activeTab = urlParams.get('tab');
            if (activeTab === 'conversion') {
                switchTab('conversion', true); // Skip URL update on page load
            } else if (activeTab === 'insurance') {
                switchTab('insurance', true); // Skip URL update on page load
            } else if (activeTab === 'remarks') {
                switchTab('remarks', true); // Skip URL update on page load
            } else if (activeTab === 'comparison') {
                switchTab('comparison', true); // Skip URL update on page load
            }
            
            // Show toast notification for conversion validation and processing
            {% if conversion_result %}
            var result = `{{ conversion_result | safe }}`;
            if (result.includes('processing completed successfully')) {
                showToast('✅ Processing completed! Formatted Insurance column added to all sheets.', false);
            } else if (result.includes('Validation successful')) {
                showToast('✅ Validation successful! Insurance Note column found in all sheets.', false);
            } else if (result.includes('Validation Error') || result.includes('❌ Error')) {
                var errorMsg = result.split('\\n')[0];
                showToast(errorMsg, true);
            }
            {% endif %}
        }
        
        // Helper function to switch tab programmatically
        function switchTabProgrammatically(tabName) {
            // Hide all tab contents
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });
            
            // Remove active class from all tabs
            document.querySelectorAll('.tab').forEach(tab => {
                tab.classList.remove('active');
            });
            
            // Show selected tab content
            document.getElementById(tabName + '-tab').classList.add('active');
            
            // Add active class to corresponding tab button
            const tabs = document.querySelectorAll('.tab');
            tabs.forEach(tab => {
                if (tab.textContent.includes(tabName === 'comparison' ? 'Comparison' : 'Conversion')) {
                    tab.classList.add('active');
                }
            });
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

def normalize_patient_id(patient_id_val):
    """Normalize patient ID for consistent matching - handles numeric, string, whitespace, leading zeros"""
    if pd.isna(patient_id_val):
        return None
    
    # Convert to string first
    patient_id_str = str(patient_id_val).strip()
    
    if not patient_id_str:
        return None
    
    # Try to convert to number to remove leading zeros and scientific notation
    try:
        # If it's a number, convert to int (removes leading zeros and decimals)
        if '.' in patient_id_str:
            num_val = float(patient_id_str)
            # If it's a whole number, convert to int
            if num_val.is_integer():
                return str(int(num_val))
            else:
                return str(num_val)
        else:
            num_val = int(patient_id_str)
            return str(num_val)
    except (ValueError, OverflowError):
        # If it's not a number, return the cleaned string
        return patient_id_str.strip()

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
            # Normalize patient ID for consistent matching
            patient_id = normalize_patient_id(patient_id_val)
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
        unmatched_patids = []  # Track unmatched PATIDs for debugging
        for idx, row in result_df.iterrows():
            patid_val = row[previous_patient_col]
            # Normalize patient ID for consistent matching
            patid = normalize_patient_id(patid_val)
            if patid and patid in appointment_insurance_map:
                # Match found - Patient ID from Appointment Report matches PATID in Smart Assist
                # Copy insurance data FROM Appointment Report TO Smart Assist file
                # Format the insurance names before adding
                insurance_data = appointment_insurance_map[patid]
                result_df.at[idx, primary_col_name] = format_insurance_name(insurance_data['primary'])
                result_df.at[idx, secondary_col_name] = format_insurance_name(insurance_data['secondary'])
                matched_count += 1
            elif patid:
                # Track unmatched PATIDs (limit to first 10 for display)
                if len(unmatched_patids) < 10:
                    unmatched_patids.append(str(patid_val))
        
        # Count statistics
        total_patients = len(result_df)
        not_matched_patients = total_patients - matched_count
        
        # Generate result message
        columns_info = ""
        if columns_added:
            columns_info = f"\n⚠️ Note: The following columns were added to Smart Assist file (with empty values): {', '.join(columns_added)}"
        
        # Add debugging info for unmatched records
        unmatched_info = ""
        if unmatched_patids and not_matched_patients > 0:
            unmatched_info = f"\n\n🔍 Sample unmatched PATIDs from Smart Assist (first 10): {', '.join(unmatched_patids)}"
            unmatched_info += f"\n💡 Tip: Check if these Patient IDs exist in the Appointment Report file with the same values."
        
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
- Successful matches: {matched_count}{unmatched_info}

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
    global conversion_data, conversion_filename, conversion_result
    global insurance_formatting_data, insurance_formatting_filename, insurance_formatting_result, insurance_formatting_output
    global remarks_appointments_data, remarks_excel_data, remarks_appointments_filename, remarks_remarks_filename, remarks_result, remarks_updated_count
    
    # Get the active tab from URL parameter
    active_tab = request.args.get('tab', 'comparison')
    
    return render_template_string(HTML_TEMPLATE, 
                                raw_data=raw_data, 
                                previous_data=previous_data,
                                raw_filename=raw_filename,
                                previous_filename=previous_filename,
                                comparison_result=comparison_result,
                                conversion_data=conversion_data,
                                conversion_filename=conversion_filename,
                                conversion_result=conversion_result,
                                insurance_formatting_data=insurance_formatting_data,
                                insurance_formatting_filename=insurance_formatting_filename,
                                insurance_formatting_result=insurance_formatting_result,
                                insurance_formatting_output=insurance_formatting_output,
                                remarks_appointments_data=remarks_appointments_data,
                                remarks_excel_data=remarks_excel_data,
                                remarks_appointments_filename=remarks_appointments_filename,
                                remarks_remarks_filename=remarks_remarks_filename,
                                remarks_result=remarks_result,
                                remarks_updated_count=remarks_updated_count,
                                active_tab=active_tab)

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
        
        # Remove "Unnamed:" columns from all sheets
        cleaned_data = {}
        for sheet_name, df in raw_data.items():
            # Convert column names to strings first, then remove columns that start with "Unnamed:"
            df.columns = df.columns.astype(str)
            df_cleaned = df.loc[:, ~df.columns.str.contains('^Unnamed:', na=False, regex=True)]
            cleaned_data[sheet_name] = df_cleaned
        raw_data = cleaned_data
        
        raw_filename = filename
        
        comparison_result = f"✅ Appointment Report uploaded successfully! Loaded {len(raw_data)} sheets: {', '.join(list(raw_data.keys()))}"
        return redirect('/comparison?tab=comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error uploading Appointment Report: {str(e)}"
        return redirect('/comparison?tab=comparison')

@app.route('/upload_previous', methods=['POST'])
def upload_previous_file():
    global previous_data, previous_filename, comparison_result
    
    if 'file' not in request.files:
        comparison_result = "❌ Error: No file provided"
        return redirect('/comparison?tab=comparison')
    
    file = request.files['file']
    if file.filename == '':
        comparison_result = "❌ Error: No file selected"
        return redirect('/comparison?tab=comparison')
    
    try:
        # Get filename without saving to disk
        filename = secure_filename(file.filename)
        
        # Read Excel file directly from memory (no disk storage)
        file.seek(0)  # Reset file pointer to beginning
        previous_data = pd.read_excel(file, sheet_name=None, engine='openpyxl')
        
        # Remove "Unnamed:" columns from all sheets
        cleaned_data = {}
        for sheet_name, df in previous_data.items():
            # Convert column names to strings first, then remove columns that start with "Unnamed:"
            df.columns = df.columns.astype(str)
            df_cleaned = df.loc[:, ~df.columns.str.contains('^Unnamed:', na=False, regex=True)]
            cleaned_data[sheet_name] = df_cleaned
        previous_data = cleaned_data
        
        previous_filename = filename
        
        comparison_result = f"✅ Smart Assist file uploaded successfully! Loaded {len(previous_data)} sheets: {', '.join(list(previous_data.keys()))}"
        return redirect('/comparison?tab=comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error uploading Smart Assist file: {str(e)}"
        return redirect('/comparison?tab=comparison')

@app.route('/compare', methods=['POST'])
def compare_files():
    global raw_data, previous_data, comparison_result
    
    if not raw_data or not previous_data:
        comparison_result = "❌ Error: Please upload both files first"
        return redirect('/comparison?tab=comparison')
    
    raw_sheet = request.form.get('raw_sheet')
    previous_sheet = request.form.get('previous_sheet')
    
    if not raw_sheet or not previous_sheet:
        comparison_result = "❌ Error: Please select sheets for both files"
        return redirect('/comparison?tab=comparison')
    
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
        
        return redirect('/comparison?tab=comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error comparing files: {str(e)}"
        return redirect('/comparison?tab=comparison')

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

@app.route('/upload_conversion', methods=['POST'])
def upload_conversion_file():
    global conversion_data, conversion_filename, conversion_result
    
    if 'file' not in request.files:
        conversion_result = "❌ Error: No file provided"
        return redirect('/comparison?tab=conversion')
    
    file = request.files['file']
    if file.filename == '':
        conversion_result = "❌ Error: No file selected"
        return redirect('/comparison?tab=conversion')
    
    try:
        # Get filename without saving to disk
        filename = secure_filename(file.filename)
        
        # Read Excel file directly from memory (no disk storage)
        file.seek(0)  # Reset file pointer to beginning
        conversion_data = pd.read_excel(file, sheet_name=None, engine='openpyxl')
        conversion_filename = filename
        
        # Remove "Unnamed:" columns from all sheets
        cleaned_data = {}
        for sheet_name, df in conversion_data.items():
            # Convert column names to strings first, then remove columns that start with "Unnamed:"
            df.columns = df.columns.astype(str)
            df_cleaned = df.loc[:, ~df.columns.str.contains('^Unnamed:', na=False, regex=True)]
            cleaned_data[sheet_name] = df_cleaned
        conversion_data = cleaned_data
        
        # Validate "Insurance Note" column exists in all sheets
        missing_sheets = []
        for sheet_name, df in conversion_data.items():
            columns = [col.lower().strip() for col in df.columns]
            if 'insurance note' not in columns:
                missing_sheets.append(sheet_name)
        
        if missing_sheets:
            conversion_result = f"❌ Validation Error: 'Insurance Note' column not found in the following sheets: {', '.join(missing_sheets)}\n\nAvailable columns in first sheet: {list(conversion_data[list(conversion_data.keys())[0]].columns) if conversion_data else 'N/A'}"
            conversion_data = None
            conversion_filename = None
        else:
            # Validation successful - now process the Insurance Note column
            processed_sheets = {}
            total_rows_processed = 0
            
            for sheet_name, df in conversion_data.items():
                # Find Insurance Note column (case-insensitive)
                insurance_note_col = None
                for col in df.columns:
                    if col.lower().strip() == 'insurance note':
                        insurance_note_col = col
                        break
                
                if insurance_note_col:
                    # Create a copy of the dataframe
                    processed_df = df.copy()
                    
                    # Extract insurance names and format them
                    def extract_and_format_insurance(note_text):
                        """Extract insurance name from Insurance Note text and format it"""
                        if pd.isna(note_text) or note_text == '':
                            return ''
                        
                        note_str = str(note_text).strip()
                        
                        # Pattern 1: "from conversion carrier: <insurance_name>| ** | ..."
                        # Pattern 2: "From Conversion Carrier: <insurance_name> | ..."
                        if 'from conversion carrier:' in note_str.lower():
                            # Find the insurance name part
                            # Look for "From Conversion Carrier:" or "from conversion carrier:"
                            parts = note_str.split('|')
                            if len(parts) > 0:
                                carrier_part = parts[0].strip()
                                # Remove "from conversion carrier:" prefix (case-insensitive)
                                if 'from conversion carrier:' in carrier_part.lower():
                                    insurance_name = carrier_part.split(':', 1)[1].strip() if ':' in carrier_part else carrier_part
                                    # Format the insurance name using existing function
                                    return format_insurance_name(insurance_name)
                        
                        # If pattern doesn't match, try to format the whole text
                        return format_insurance_name(note_str)
                    
                    # Extract status from Insurance Note
                    def extract_status(note_text):
                        """Extract status from Insurance Note text"""
                        if pd.isna(note_text) or note_text == '':
                            return ''
                        
                        note_str = str(note_text).strip()
                        
                        # Pattern: "Status - <status>"
                        # Look for "Status -" or "Status-"
                        import re
                        status_pattern = r'Status\s*-\s*([^\]|]+)'
                        match = re.search(status_pattern, note_str, re.IGNORECASE)
                        if match:
                            status = match.group(1).strip()
                            # Remove any trailing brackets or pipes
                            status = status.rstrip(' ]|')
                            return status
                        
                        return ''
                    
                    # Apply extraction and formatting
                    processed_df['Formatted Insurance'] = processed_df[insurance_note_col].apply(extract_and_format_insurance)
                    processed_df['Status'] = processed_df[insurance_note_col].apply(extract_status)
                    processed_df['Conversion'] = 'Conversion'  # Add Conversion column with value "Conversion"
                    
                    # Find position of Insurance Note column and insert new columns after it
                    insurance_note_index = processed_df.columns.get_loc(insurance_note_col)
                    # Move the new columns to right after Insurance Note
                    cols = processed_df.columns.tolist()
                    # Remove new columns from their current position
                    for col_name in ['Formatted Insurance', 'Status', 'Conversion']:
                        if col_name in cols:
                            cols.remove(col_name)
                    # Insert new columns after Insurance Note
                    cols.insert(insurance_note_index + 1, 'Formatted Insurance')
                    cols.insert(insurance_note_index + 2, 'Status')
                    cols.insert(insurance_note_index + 3, 'Conversion')
                    processed_df = processed_df[cols]
                    
                    processed_sheets[sheet_name] = processed_df
                    total_rows_processed += len(processed_df)
                else:
                    processed_sheets[sheet_name] = df
            
            # Update conversion_data with processed data
            conversion_data = processed_sheets
            
            conversion_result = f"✅ Validation and processing completed successfully!\n\n📊 File loaded: {filename}\n📋 Sheets processed: {len(conversion_data)}\n📋 Sheet names: {', '.join(list(conversion_data.keys()))}\n📊 Total rows processed: {total_rows_processed}\n\n✅ New columns added:\n- 'Formatted Insurance' - Extracted and formatted insurance names\n- 'Status' - Extracted status values\n- 'Conversion' - Set to 'Conversion' for all rows\n💾 Ready to download the processed file!"
        
        return redirect('/comparison?tab=conversion')
        
    except Exception as e:
        conversion_result = f"❌ Error uploading Conversion Report: {str(e)}"
        conversion_data = None
        conversion_filename = None
        return redirect('/comparison?tab=conversion')

@app.route('/download_conversion', methods=['POST'])
def download_conversion_result():
    global conversion_data, conversion_filename
    
    if not conversion_data:
        return jsonify({'error': 'No data to download'}), 400
    
    filename = request.form.get('filename', '').strip()
    if not filename:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"conversion_report_{timestamp}.xlsx"
    
    try:
        # Create a temporary file
        import tempfile
        temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
        
        try:
            with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                for sheet_name, df in conversion_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Clear data after successful download
            global conversion_result
            conversion_data = None
            conversion_filename = None
            conversion_result = None
            
            return send_file(temp_path, as_attachment=True, download_name=filename)
            
        finally:
            # Clean up temporary file
            os.close(temp_fd)
            if os.path.exists(temp_path):
                os.unlink(temp_path)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/reset_comparison', methods=['POST'])
def reset_comparison():
    global raw_data, previous_data, raw_filename, previous_filename, comparison_result
    # Explicitly do NOT touch conversion_data, conversion_filename, or conversion_result
    
    try:
        # Reset ONLY comparison tool variables - do not affect conversion tool
        raw_data = {}
        previous_data = {}
        raw_filename = None
        previous_filename = None
        comparison_result = "🔄 Comparison tool reset successfully! All files and data have been cleared."
        
        return redirect('/comparison?tab=comparison')
        
    except Exception as e:
        comparison_result = f"❌ Error resetting comparison tool: {str(e)}"
        return redirect('/comparison?tab=comparison')

@app.route('/reset_conversion', methods=['POST'])
def reset_conversion():
    global conversion_data, conversion_filename, conversion_result
    # Explicitly do NOT touch raw_data, previous_data, raw_filename, previous_filename, or comparison_result
    
    try:
        # Reset ONLY conversion tool variables - do not affect comparison tool
        conversion_data = {}
        conversion_filename = None
        conversion_result = "🔄 Conversion tool reset successfully! All files and data have been cleared."
        
        return redirect('/comparison?tab=conversion')
        
    except Exception as e:
        conversion_result = f"❌ Error resetting conversion tool: {str(e)}"
        return redirect('/comparison?tab=conversion')

def reformat_insurance_column(df, source_column, new_column_name):
    """Reformat the source column and create a new column next to it"""
    if source_column not in df.columns:
        return None, f"Column '{source_column}' not found in sheet."
    
    # Apply the reformatting
    df[new_column_name] = df[source_column].apply(format_insurance_name)
    
    # Get the position of the source column
    source_col_idx = df.columns.get_loc(source_column)
    
    # Reorder columns to place new column right after source column
    cols = df.columns.tolist()
    # Remove new column from its current position
    cols.remove(new_column_name)
    # Insert it right after source column
    cols.insert(source_col_idx + 1, new_column_name)
    df = df[cols]
    
    return df, None

def process_insurance_formatting(data_dict, source_column="Dental Primary Ins Carr", new_column_name="formated insurance names"):
    """Process all sheets in the Excel file for insurance formatting"""
    output_lines = []
    output_lines.append("=" * 70)
    output_lines.append("PROCESSING EXCEL FILE - REFORMATTING COLUMNS")
    output_lines.append("=" * 70)
    output_lines.append("")
    
    processed_sheets = {}
    
    for sheet_name, df in data_dict.items():
        output_lines.append(f"📋 Processing sheet: {sheet_name}")
        output_lines.append(f"   Original shape: {df.shape[0]} rows × {df.shape[1]} columns")
        
        if source_column not in df.columns:
            output_lines.append(f"   ⚠️  Column '{source_column}' not found. Skipping this sheet.")
            processed_sheets[sheet_name] = df
            output_lines.append("")
            continue
        
        # Reformat the column
        df_processed, error = reformat_insurance_column(df.copy(), source_column, new_column_name)
        
        if error:
            output_lines.append(f"   ❌ Error: {error}")
            processed_sheets[sheet_name] = df
        else:
            processed_sheets[sheet_name] = df_processed
            output_lines.append(f"   ✅ Column reformatted successfully!")
            output_lines.append(f"   New shape: {df_processed.shape[0]} rows × {df_processed.shape[1]} columns")
            
            # Show sample
            sample_df = df_processed[[source_column, new_column_name]].head(10)
            output_lines.append(f"   Sample of original vs formatted (first 10 rows):")
            output_lines.append(sample_df.to_string(index=False))
            output_lines.append(f"   Total formatted entries: {df_processed[new_column_name].notna().sum()}")
            output_lines.append(f"   Unique formatted values: {df_processed[new_column_name].value_counts().head(10).to_string()}")
        
        output_lines.append("")
    
    output_lines.append("=" * 70)
    output_lines.append("PROCESSING COMPLETE!")
    output_lines.append("=" * 70)
    
    return processed_sheets, "\n".join(output_lines)

@app.route('/upload_insurance_formatting', methods=['POST'])
def upload_insurance_formatting():
    global insurance_formatting_data, insurance_formatting_filename, insurance_formatting_result, insurance_formatting_output
    
    if 'file' not in request.files:
        insurance_formatting_result = "❌ Error: No file provided"
        return redirect('/comparison?tab=insurance')
    
    file = request.files['file']
    if file.filename == '':
        insurance_formatting_result = "❌ Error: No file selected"
        return redirect('/comparison?tab=insurance')
    
    try:
        # Get filename without saving to disk
        filename = secure_filename(file.filename)
        
        # Read Excel file directly from memory (no disk storage)
        file.seek(0)  # Reset file pointer to beginning
        excel_data = pd.read_excel(file, sheet_name=None, engine='openpyxl')
        
        # Remove "Unnamed:" columns from all sheets
        cleaned_data = {}
        for sheet_name, df in excel_data.items():
            # Convert column names to strings first, then remove columns that start with "Unnamed:"
            df.columns = df.columns.astype(str)
            df_cleaned = df.loc[:, ~df.columns.str.contains('^Unnamed:', na=False, regex=True)]
            cleaned_data[sheet_name] = df_cleaned
        
        # Process all sheets automatically
        insurance_formatting_data, insurance_formatting_output = process_insurance_formatting(cleaned_data)
        insurance_formatting_filename = filename
        
        # Count sheets processed
        sheets_count = len(insurance_formatting_data)
        insurance_formatting_result = f"✅ Processing complete! Processed {sheets_count} sheet(s). Formatted insurance names column added to all sheets."
        
        return redirect('/comparison?tab=insurance')
        
    except Exception as e:
        insurance_formatting_result = f"❌ Error processing file: {str(e)}"
        insurance_formatting_output = f"Error: {str(e)}"
        return redirect('/comparison?tab=insurance')

@app.route('/download_insurance_formatting', methods=['POST'])
def download_insurance_formatting():
    global insurance_formatting_data, insurance_formatting_filename, insurance_formatting_result, insurance_formatting_output
    
    if not insurance_formatting_data:
        return jsonify({'error': 'No data to download'}), 400
    
    filename = request.form.get('filename', '').strip()
    if not filename:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"formatted_insurance_names_{timestamp}.xlsx"
    
    try:
        # Create a temporary file
        import tempfile
        temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
        
        try:
            with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                for sheet_name, df in insurance_formatting_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Clear data after successful download
            insurance_formatting_data = None
            insurance_formatting_filename = None
            insurance_formatting_result = None
            insurance_formatting_output = ""
            
            return send_file(temp_path, as_attachment=True, download_name=filename)
            
        finally:
            # Clean up temporary file
            os.close(temp_fd)
            if os.path.exists(temp_path):
                os.unlink(temp_path)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/reset_insurance_formatting', methods=['POST'])
def reset_insurance_formatting():
    global insurance_formatting_data, insurance_formatting_filename, insurance_formatting_result, insurance_formatting_output
    # Explicitly do NOT touch other tool variables
    
    try:
        # Reset ONLY insurance formatting tool variables
        insurance_formatting_data = {}
        insurance_formatting_filename = None
        insurance_formatting_result = "🔄 Insurance formatting tool reset successfully! All files and data have been cleared."
        insurance_formatting_output = ""
        
        return redirect('/comparison?tab=insurance')
        
    except Exception as e:
        insurance_formatting_result = f"❌ Error resetting insurance formatting tool: {str(e)}"
        return redirect('/comparison?tab=insurance')

# Remarks update functions (from excel-remark)
def process_remarks_excel_file(file_stream):
    """Process uploaded Excel file and extract Patient ID, Remark, and Agent Name data.
    Checks all sheets to find the one with required columns."""
    try:
        from openpyxl import load_workbook
        
        wb = load_workbook(file_stream)
        ws = None
        patient_id_col = None
        remark_col = None
        agent_name_col = None
        header_row = 1
        
        # Try all sheets to find the one with Patient ID, Remark, and Agent Name columns
        for sheet_name in wb.sheetnames:
            current_ws = wb[sheet_name]
            
            # Look for headers in the first few rows
            for row_num in range(1, min(6, current_ws.max_row + 1)):
                temp_patient_id_col = None
                temp_remark_col = None
                temp_agent_name_col = None
                
                for col in range(1, current_ws.max_column + 1):
                    cell_value = str(current_ws.cell(row=row_num, column=col).value or '').strip().lower()
                    cell_value_clean = cell_value.replace(' ', '').replace('_', '').replace('-', '').replace('.', '')
                    
                    if ('patient' in cell_value_clean and 'id' in cell_value_clean) or cell_value_clean == 'pid':
                        temp_patient_id_col = col
                    elif 'remark' in cell_value_clean:
                        temp_remark_col = col
                    elif 'agent' in cell_value_clean and 'name' in cell_value_clean:
                        temp_agent_name_col = col
                
                # If we found Patient ID and Remark (Agent Name is optional), use this sheet
                if temp_patient_id_col and temp_remark_col:
                    patient_id_col = temp_patient_id_col
                    remark_col = temp_remark_col
                    agent_name_col = temp_agent_name_col
                    header_row = row_num
                    ws = current_ws
                    break
            
            if ws:
                break
        
        # If still not found, try active sheet
        if not ws:
            ws = wb.active
            for col in range(1, ws.max_column + 1):
                cell_value = str(ws.cell(row=1, column=col).value or '').strip().lower()
                cell_value_clean = cell_value.replace(' ', '').replace('_', '').replace('-', '').replace('.', '')
                
                if ('patient' in cell_value_clean and 'id' in cell_value_clean) or cell_value_clean == 'pid':
                    patient_id_col = col
                elif 'remark' in cell_value_clean:
                    remark_col = col
                elif 'agent' in cell_value_clean and 'name' in cell_value_clean:
                    agent_name_col = col
        
        if not patient_id_col:
            sheet_names = ', '.join(wb.sheetnames)
            raise Exception(f"Patient ID column not found in Excel file. Checked sheets: {sheet_names}")
        
        if not remark_col:
            sheet_names = ', '.join(wb.sheetnames)
            raise Exception(f"Remark column not found in Excel file. Checked sheets: {sheet_names}")
        
        # Extract data - now returns list of records for each patient ID
        excel_data = {}
        data_start_row = header_row + 1
        for row in range(data_start_row, ws.max_row + 1):  # Skip header row
            patient_id = str(ws.cell(row=row, column=patient_id_col).value or '').strip()
            remark = str(ws.cell(row=row, column=remark_col).value or '').strip()
            agent_name = str(ws.cell(row=row, column=agent_name_col).value or '').strip() if agent_name_col else ''
            
            # Clean up Patient ID - remove .0 if it's a float
            if patient_id.endswith('.0'):
                patient_id = patient_id[:-2]
            
            if patient_id:  # Only add non-empty patient IDs
                if patient_id not in excel_data:
                    excel_data[patient_id] = []
                
                excel_data[patient_id].append({
                    'remark': remark,
                    'agent_name': agent_name
                })
        
        return excel_data
        
    except Exception as e:
        raise Exception(f"Error processing Excel file: {str(e)}")

def process_remarks_appointments_excel(file_stream):
    """Read an appointments Excel and return list of appointment dicts with all columns.
    
    Only requires 'Pat ID' column. All other columns are preserved as-is.
    Checks all sheets to find the one with Pat ID column.
    """
    from openpyxl import load_workbook

    wb = load_workbook(file_stream)
    ws = None
    headers = []
    pat_id_col = None
    header_row = 1
    
    # Try all sheets to find the one with Pat ID column
    for sheet_name in wb.sheetnames:
        current_ws = wb[sheet_name]
        
        # Try to find header row (check first 5 rows)
        for row_num in range(1, min(6, current_ws.max_row + 1)):
            temp_headers = []
            for col in range(1, current_ws.max_column + 1):
                raw = current_ws.cell(row=row_num, column=col).value
                name = (str(raw or '')).strip()
                temp_headers.append(name)
            
            # Check if this row looks like headers (has a Pat ID column)
            for i, header in enumerate(temp_headers):
                header_lower = header.lower().replace(' ', '').replace('_', '').replace('-', '').replace('.', '')
                # Check for various patterns: pat id, patient id, patientid, patid, etc.
                if ('pat' in header_lower and 'id' in header_lower) or header_lower == 'pid':
                    headers = temp_headers
                    header_row = row_num
                    pat_id_col = i + 1  # 1-based column index
                    ws = current_ws
                    break
            
            if pat_id_col:
                break
        
        if pat_id_col:
            break
    
    # If still not found, use active sheet and check again
    if pat_id_col is None:
        ws = wb.active
        headers = []
        for col in range(1, ws.max_column + 1):
            raw = ws.cell(row=1, column=col).value
            name = (str(raw or '')).strip()
            headers.append(name)
        
        for i, header in enumerate(headers):
            header_lower = header.lower().replace(' ', '').replace('_', '').replace('-', '').replace('.', '')
            if ('pat' in header_lower and 'id' in header_lower) or header_lower == 'pid':
                pat_id_col = i + 1
                header_row = 1
                break
    
    if pat_id_col is None:
        # Provide helpful error message with found columns
        sheet_names = ', '.join(wb.sheetnames)
        found_columns = ', '.join([f"'{h}'" for h in headers if h]) or 'none'
        raise Exception(f"Pat ID column not found in appointments Excel. Checked sheets: {sheet_names}. Found columns: {found_columns}. Please ensure there's a column containing 'Pat ID', 'Patient ID', or similar.")

    # Read all rows starting after the header row
    appointments = []
    data_start_row = header_row + 1
    for row in range(data_start_row, ws.max_row + 1):
        record = {}
        
        # Read all columns
        for col, header in enumerate(headers, 1):
            value = ws.cell(row=row, column=col).value
            record[header] = '' if value is None else str(value)
        
        # Normalize Patient ID to string without trailing .0
        pat_id_value = record.get(headers[pat_id_col - 1], '')
        pid = str(pat_id_value).strip()
        if pid.endswith('.0'):
            pid = pid[:-2]
        record['Pat ID'] = pid  # Standardize the key name
        
        # Ensure Remark and Agent Name exist
        if 'Remark' not in record:
            record['Remark'] = ''
        if 'Agent Name' not in record:
            record['Agent Name'] = ''
        
        # Skip empty rows (no Pat ID)
        if pid:
            appointments.append(record)

    return appointments

def update_appointments_with_remarks(appointments, excel_data):
    """Update appointments with remarks and agent names from Excel data based on Patient ID matching.
    Creates separate rows for each match when Patient ID appears multiple times."""
    updated_appointments = []
    updated_count = 0
    
    for appointment in appointments:
        patient_id = str(appointment.get('Pat ID', '')).strip()
        matches_found = False
        
        # Try exact match first
        if patient_id and patient_id in excel_data:
            # Create a separate row for each match
            for match_data in excel_data[patient_id]:
                new_appointment = appointment.copy()  # Copy all original data
                new_appointment['Remark'] = match_data['remark']
                new_appointment['Agent Name'] = match_data['agent_name']
                updated_appointments.append(new_appointment)
                updated_count += 1
                matches_found = True
        # Try with .0 suffix (in case Excel has float format)
        elif patient_id and f"{patient_id}.0" in excel_data:
            for match_data in excel_data[f"{patient_id}.0"]:
                new_appointment = appointment.copy()  # Copy all original data
                new_appointment['Remark'] = match_data['remark']
                new_appointment['Agent Name'] = match_data['agent_name']
                updated_appointments.append(new_appointment)
                updated_count += 1
                matches_found = True
        
        # If no matches found, add original appointment with empty remark and agent name
        if not matches_found:
            appointment['Remark'] = ''
            appointment['Agent Name'] = ''
            updated_appointments.append(appointment)
    
    return updated_appointments, updated_count

def create_excel_from_appointments(appointments, filename):
    """Create Excel file from processed appointment data with all columns."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    import io
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Appointment Data"
    
    if not appointments:
        return wb
    
    # Get all unique headers from all appointments
    all_headers = set()
    for appointment in appointments:
        all_headers.update(appointment.keys())
    
    # Convert to list and ensure Pat ID, Insurance Name, Remark, and Agent Name are at the end for visibility
    headers = list(all_headers)
    if 'Pat ID' in headers:
        headers.remove('Pat ID')
    if 'Insurance Name' in headers:
        headers.remove('Insurance Name')
    if 'Remark' in headers:
        headers.remove('Remark')
    if 'Agent Name' in headers:
        headers.remove('Agent Name')
    headers.extend(['Pat ID', 'Insurance Name', 'Remark', 'Agent Name'])  # Put these at the end
    
    # Set headers
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    # Add appointment data
    for row, appointment in enumerate(appointments, 2):
        for col, header in enumerate(headers, 1):
            value = appointment.get(header, '')
            ws.cell(row=row, column=col, value=value)
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Save to memory
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    
    return excel_buffer

@app.route('/upload_remarks', methods=['POST'])
def upload_remarks():
    global remarks_appointments_data, remarks_excel_data, remarks_appointments_filename, remarks_remarks_filename, remarks_result, remarks_updated_count
    
    appointments_file = request.files.get('appointments_file')
    remarks_file = request.files.get('remarks_file')
    
    # Require both Excel files
    if not appointments_file or appointments_file.filename == '':
        remarks_result = "❌ Error: Please upload the Appointments Excel file."
        return redirect('/comparison?tab=remarks')
    
    if not remarks_file or remarks_file.filename == '':
        remarks_result = "❌ Error: Please upload the Remarks Excel file."
        return redirect('/comparison?tab=remarks')
    
    try:
        # Process appointments Excel directly from memory
        appointments_filename_raw = secure_filename(appointments_file.filename)
        appointments_file.seek(0)  # Reset file pointer
        remarks_appointments_data = process_remarks_appointments_excel(appointments_file)
        remarks_appointments_filename = appointments_filename_raw
        
        # Process remarks Excel
        remarks_file.seek(0)  # Reset file pointer
        remarks_excel_data = process_remarks_excel_file(remarks_file)
        remarks_remarks_filename = secure_filename(remarks_file.filename)
        
        # Update appointments with remarks
        updated_appointments, updated_count = update_appointments_with_remarks(remarks_appointments_data, remarks_excel_data)
        
        # Format Insurance Note column to create Insurance Name column
        for appointment in updated_appointments:
            # Find Insurance Note column (case-insensitive search)
            insurance_note_value = None
            insurance_note_key = None
            
            for key, value in appointment.items():
                if key.lower().strip() == 'insurance note':
                    insurance_note_key = key
                    insurance_note_value = value
                    break
            
            # Format the insurance note value using existing format_insurance_name function
            if insurance_note_value:
                # Handle None, empty string, or NaN values
                insurance_str = str(insurance_note_value).strip() if insurance_note_value else ''
                if insurance_str and insurance_str.lower() not in ['nan', 'none', '']:
                    try:
                        # Extract insurance name from "from conversion carrier:" pattern (case-insensitive)
                        # Example 1: "from conversion carrier: <insurance_name> | Plan #684 | ..."
                        # Example 2: "2025-10-29 18:34:01 PT, By - 5240HSHINDE, Status - Eligible ] From conversion carrier: <insurance_name>"
                        # We want to extract only the insurance name after "from conversion carrier:" and before the first "|" (if present)
                        extracted_insurance_name = None
                        
                        # Check for "from conversion carrier:" pattern (case-insensitive)
                        if 'from conversion carrier:' in insurance_str.lower():
                            # Find the position of "from conversion carrier:" (case-insensitive)
                            import re
                            pattern_match = re.search(r'from conversion carrier:\s*(.+?)(?:\s*\||$)', insurance_str, re.IGNORECASE)
                            
                            if pattern_match:
                                # Extract the insurance name (everything after "from conversion carrier:" until "|" or end of string)
                                insurance_name = pattern_match.group(1).strip()
                                # Remove any trailing characters like "]" if present
                                insurance_name = insurance_name.rstrip(' ]|')
                                extracted_insurance_name = insurance_name
                            else:
                                # Fallback: Split by "|" to get only the part before the first pipe
                                parts = insurance_str.split('|', 1)  # Split only on first "|"
                                if len(parts) > 0:
                                    carrier_part = parts[0].strip()
                                    # Find "from conversion carrier:" in this part (case-insensitive)
                                    carrier_match = re.search(r'from conversion carrier:\s*(.+?)$', carrier_part, re.IGNORECASE)
                                    if carrier_match:
                                        insurance_name = carrier_match.group(1).strip()
                                        insurance_name = insurance_name.rstrip(' ]|')
                                        extracted_insurance_name = insurance_name
                                    elif 'from conversion carrier:' in carrier_part.lower():
                                        # Extract everything after "from conversion carrier:"
                                        insurance_name = carrier_part.split(':', 1)[1].strip() if ':' in carrier_part else carrier_part
                                        insurance_name = insurance_name.rstrip(' ]|')
                                        extracted_insurance_name = insurance_name
                        
                        # If we extracted a name, use it; otherwise use the full text
                        insurance_to_format = extracted_insurance_name if extracted_insurance_name else insurance_str
                        
                        # Format the insurance name using existing function
                        formatted_insurance = format_insurance_name(insurance_to_format)
                        appointment['Insurance Name'] = formatted_insurance if formatted_insurance and not (isinstance(formatted_insurance, float) and pd.isna(formatted_insurance)) else ''
                    except Exception as e:
                        # If formatting fails, use empty string
                        appointment['Insurance Name'] = ''
                else:
                    appointment['Insurance Name'] = ''
            else:
                appointment['Insurance Name'] = ''
        
        # Update the global processed_appointments with the new data
        remarks_appointments_data = updated_appointments
        remarks_updated_count = updated_count
        
        remarks_result = f"✅ Successfully processed {len(remarks_appointments_data)} appointment(s) and updated {updated_count} appointment(s) with remarks and agent names. Insurance Name column added based on Insurance Note formatting."
        
        return redirect('/comparison?tab=remarks')
        
    except Exception as e:
        remarks_result = f"❌ Error processing files: {str(e)}"
        return redirect('/comparison?tab=remarks')

@app.route('/download_remarks', methods=['POST'])
def download_remarks():
    global remarks_appointments_data, remarks_appointments_filename, remarks_remarks_filename, remarks_result, remarks_updated_count
    
    if not remarks_appointments_data:
        return jsonify({'error': 'No data to download'}), 400
    
    filename = request.form.get('filename', '').strip()
    if not filename:
        if remarks_appointments_filename:
            base_name = os.path.splitext(remarks_appointments_filename)[0]
            filename = f"{base_name}_appointments.xlsx"
        else:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"appointments_with_remarks_{timestamp}.xlsx"
    
    try:
        # Create Excel file
        excel_buffer = create_excel_from_appointments(remarks_appointments_data, filename)
        
        # Clear data after successful download
        remarks_appointments_data = None
        remarks_excel_data = None
        remarks_appointments_filename = None
        remarks_remarks_filename = None
        remarks_result = None
        remarks_updated_count = 0
        
        return send_file(excel_buffer, 
                        as_attachment=True, 
                        download_name=filename,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/reset_remarks', methods=['POST'])
def reset_remarks():
    global remarks_appointments_data, remarks_excel_data, remarks_appointments_filename, remarks_remarks_filename, remarks_result, remarks_updated_count
    # Explicitly do NOT touch other tool variables
    
    try:
        # Reset ONLY remarks tool variables
        remarks_appointments_data = None
        remarks_excel_data = None
        remarks_appointments_filename = None
        remarks_remarks_filename = None
        remarks_result = "🔄 Remarks tool reset successfully! All files and data have been cleared."
        remarks_updated_count = 0
        
        return redirect('/comparison?tab=remarks')
        
    except Exception as e:
        remarks_result = f"❌ Error resetting remarks tool: {str(e)}"
        return redirect('/comparison?tab=remarks')

@app.route('/reset_app', methods=['POST'])
def reset_app():
    global raw_data, previous_data, raw_filename, previous_filename, comparison_result
    global conversion_data, conversion_filename, conversion_result
    global insurance_formatting_data, insurance_formatting_filename, insurance_formatting_result, insurance_formatting_output
    global remarks_appointments_data, remarks_excel_data, remarks_appointments_filename, remarks_remarks_filename, remarks_result, remarks_updated_count
    
    try:
        # Reset all global variables
        raw_data = {}
        previous_data = {}
        raw_filename = None
        previous_filename = None
        comparison_result = "🔄 Application reset successfully! All files and data have been cleared."
        
        # Reset conversion data
        conversion_data = {}
        conversion_filename = None
        conversion_result = None
        
        # Reset insurance formatting data
        insurance_formatting_data = {}
        insurance_formatting_filename = None
        insurance_formatting_result = None
        insurance_formatting_output = ""
        
        # Reset remarks data
        remarks_appointments_data = None
        remarks_excel_data = None
        remarks_appointments_filename = None
        remarks_remarks_filename = None
        remarks_result = None
        remarks_updated_count = 0
        
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
