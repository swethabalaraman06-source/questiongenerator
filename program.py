import json
import pandas as pd
from flask import Flask, render_template_string, request, url_for, redirect
import random
import re
import os
import pdfplumber  # Added for PDF processing
import io
import docx
from docx import Document
from flask import send_file
from docx.enum.text import WD_ALIGN_PARAGRAPH
from flask import session, send_file

app = Flask(__name__)
app.secret_key = 'your_super_secret_key_here'

# --- STEP 1: UPLOAD PAGE (Updated label to include PDF) ---
UPLOAD_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Question Generator</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
</head>
<body class="container py-5">
    <div class="card p-4 shadow-sm mx-auto" style="max-width: 500px;">
        <h2 class="text-center mb-4">Upload Question Bank</h2>
        <form action="/setup" method="post" enctype="multipart/form-data">
            <input type="file" name="file" class="form-control mb-3" required>
            <button type="submit" class="btn btn-primary w-100">Load Bank</button>
        </form>
    </div>
</body>
</html>
"""

# --- STEP 2: DYNAMIC CONFIGURATION ---
CONFIG_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Step 2: Configuration</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
        .section-header { background: #e9ecef; padding: 10px; font-weight: bold; margin: 20px 0; border-radius: 5px; }
        .table-primary { --bs-table-bg: #0d6efd; --bs-table-color: white; }
        .sticky-bottom-btn { position: sticky; bottom: 0; background: white; padding: 20px 0; border-top: 1px solid #ddd; z-index: 100; }
    </style>
</head>
<body class="container py-5">
    <div class="card p-4 shadow-sm mb-5">
        <h3 class="mb-4 text-primary text-center">Question Paper Configuration</h3>
        <form action="/generate" method="post">
           
            <div class="section-header">1. Select Units for Exam</div>
            <div class="d-flex justify-content-between mb-4 px-2">
                {% for i in range(1, 6) %}
                <div class="form-check form-check-inline">
                    <input class="form-check-input unit-check" type="checkbox" name="selected_units"
                           value="{{ i }}" id="u{{ i }}" checked onchange="filterCOs()">
                    <label class="form-check-label fw-bold" for="u{{ i }}">Unit {{ i }}</label>
                </div>
                {% endfor %}
            </div>

            <div class="section-header">PART A (5 Questions)</div>
            <div class="table-responsive mb-4">
                <table class="table table-bordered align-middle text-center">
                    <thead class="table-primary">
                        <tr><th>Q.No</th><th>CO</th><th>K-Level</th><th>Type</th></tr>
                    </thead>
                    <tbody>
                        {% for i in range(1, 6) %}
                        <tr>
                            <td class="fw-bold">Q{{ i }}</td>
                            <td><select name="pa_co{{ i }}" class="form-select co-dropdown"></select></td>
                            <td>
                                <select name="pa_k{{ i }}" class="form-select">
                                    {% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}
                                </select>
                            </td>
                            <td>
                                {% if i == 2 or i == 4 %}<span class="badge bg-danger">MCQ</span>
                                {% else %}<span class="badge bg-secondary">Theory</span>{% endif %}
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>

            <div class="section-header">PART B (6a & 6b Separate)</div>
            <div class="row g-3 mb-4 text-center">
                <div class="col-md-3"><label class="fw-bold">6(a) CO</label><select name="pb_co_a" id="pb_co_a" class="form-select co-dropdown" onchange="syncPartB()"></select></div>
                <div class="col-md-3"><label class="fw-bold">6(a) K</label><select name="pb_k_a" id="pb_k_a" class="form-select" onchange="syncPartB()">{% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}</select></div>
                <div class="col-md-3"><label class="fw-bold">6(b) CO</label><select name="pb_co_b" id="pb_co_b" class="form-select co-dropdown"></select></div>
                <div class="col-md-3"><label class="fw-bold">6(b) K</label><select name="pb_k_b" id="pb_k_b" class="form-select">{% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}</select></div>
            </div>

            <div class="section-header">PART C (Q7 & Q8 Mutually Exclusive Config)</div>
            <div class="row mb-4">
                <div class="col-md-6 border-end">
                    <div class="d-flex justify-content-between align-items-center mb-2">
                        <h5 class="text-secondary mb-0">Question 7</h5>
                        <div class="form-check form-switch">
                            <input class="form-check-input q-master-toggle" type="checkbox" id="toggle7" checked onchange="handleMasterToggle('7')">
                            <label class="form-check-label fw-bold" for="toggle7">Enable Direct 16</label>
                        </div>
                    </div>
                    <label class="fw-bold">Pattern Choice:</label>
                    <select name="q7_pattern" id="q7_pattern" class="form-select mb-3" onchange="toggleInputs('7')">
                        <option value="16">Single 16 Marks</option>
                        <option value="8+8">Split 8 + 8</option>
                        <option value="10+6">Split 10 + 6</option>
                    </select>
                    <div class="card p-2 mb-2 bg-light">
                        <h6 class="fw-bold">7 (a) Subdivision</h6>
                        <div class="row g-2">
                            <div class="col-6">
                                <label class="small">i) Unit/CO</label>
                                <select name="pc_co7ai" class="form-select co-dropdown" onchange="syncSub('pc_co7ai', 'pc_co7aii')"></select>
                                <label class="small">K-Level</label>
                                <select name="pc_k7ai" class="form-select" onchange="syncSub('pc_k7ai', 'pc_k7aii')">{% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}</select>
                            </div>
                            <div id="div_7a_ii" class="col-6 d-none">
                                <label class="small">ii) Unit/CO</label>
                                <select name="pc_co7aii" class="form-select co-dropdown"></select>
                                <label class="small">K-Level</label>
                                <select name="pc_k7aii" class="form-select">{% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}</select>
                            </div>
                        </div>
                    </div>
                    <div class="card p-2 bg-light">
                        <h6 class="fw-bold">7 (b) Subdivision</h6>
                        <div class="row g-2">
                            <div class="col-6">
                                <label class="small">i) Unit/CO</label>
                                <select name="pc_co7bi" class="form-select co-dropdown" onchange="syncSub('pc_co7bi', 'pc_co7bii')"></select>
                                <label class="small">K-Level</label>
                                <select name="pc_k7bi" class="form-select" onchange="syncSub('pc_k7bi', 'pc_k7bii')">{% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}</select>
                            </div>
                            <div id="div_7b_ii" class="col-6 d-none">
                                <label class="small">ii) Unit/CO</label>
                                <select name="pc_co7bii" class="form-select co-dropdown"></select>
                                <label class="small">K-Level</label>
                                <select name="pc_k7bii" class="form-select">{% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}</select>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-md-6">
                    <div class="d-flex justify-content-between align-items-center mb-2">
                        <h5 class="text-secondary mb-0">Question 8</h5>
                        <div class="form-check form-switch">
                            <input class="form-check-input q-master-toggle" type="checkbox" id="toggle8" onchange="handleMasterToggle('8')">
                            <label class="form-check-label fw-bold" for="toggle8">Enable Direct 16</label>
                        </div>
                    </div>
                    <label class="fw-bold">Pattern Choice:</label>
                    <select name="q8_pattern" id="q8_pattern" class="form-select mb-3" onchange="toggleInputs('8')">
                        <option value="16">Single 16 Marks</option>
                        <option value="8+8">Split 8 + 8</option>
                        <option value="10+6">Split 10 + 6</option>
                    </select>
                    <div class="card p-2 mb-2 bg-light">
                        <h6 class="fw-bold">8 (a) Subdivision</h6>
                        <div class="row g-2">
                            <div class="col-6">
                                <label class="small">i) Unit/CO</label>
                                <select name="pc_co8ai" class="form-select co-dropdown" onchange="syncSub('pc_co8ai', 'pc_co8aii')"></select>
                                <label class="small">K-Level</label>
                                <select name="pc_k8ai" class="form-select" onchange="syncSub('pc_k8ai', 'pc_k8aii')">{% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}</select>
                            </div>
                            <div id="div_8a_ii" class="col-6 d-none">
                                <label class="small">ii) Unit/CO</label>
                                <select name="pc_co8aii" class="form-select co-dropdown"></select>
                                <label class="small">K-Level</label>
                                <select name="pc_k8aii" class="form-select">{% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}</select>
                            </div>
                        </div>
                    </div>
                    <div class="card p-2 bg-light">
                        <h6 class="fw-bold">8 (b) Subdivision</h6>
                        <div class="row g-2">
                            <div class="col-6">
                                <label class="small">i) Unit/CO</label>
                                <select name="pc_co8bi" class="form-select co-dropdown" onchange="syncSub('pc_co8bi', 'pc_co8bii')"></select>
                                <label class="small">K-Level</label>
                                <select name="pc_k8bi" class="form-select" onchange="syncSub('pc_k8bi', 'pc_k8bii')">{% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}</select>
                            </div>
                            <div id="div_8b_ii" class="col-6 d-none">
                                <label class="small">ii) Unit/CO</label>
                                <select name="pc_co8bii" class="form-select co-dropdown"></select>
                                <label class="small">K-Level</label>
                                <select name="pc_k8bii" class="form-select">{% for k in k_levels %}<option value="{{ k }}">{{ k }}</option>{% endfor %}</select>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <div class="sticky-bottom-btn">
                <button type="submit" class="btn btn-success btn-lg w-100 shadow fw-bold">GENERATE QUESTION PAPER</button>
            </div>
        </form>
    </div>

    <script>
        const mapping = {{ mapping|tojson }};
        function filterCOs() {
            const selectedUnits = Array.from(document.querySelectorAll('.unit-check:checked')).map(cb => cb.value);
            const dropdowns = document.querySelectorAll('.co-dropdown');
            dropdowns.forEach(dropdown => {
                const prev = dropdown.value;
                dropdown.innerHTML = '';
                selectedUnits.forEach(unit => {
                    if (mapping[unit]) {
                        mapping[unit].forEach(co => {
                            const opt = document.createElement('option');
                            opt.value = co; opt.innerHTML = co + " (Unit " + unit + ")";
                            if (co === prev) opt.selected = true;
                            dropdown.appendChild(opt);
                        });
                    }
                });
            });
        }

        function handleMasterToggle(qNum) {
            const otherQ = (qNum === '7') ? '8' : '7';
            const currentToggle = document.getElementById(`toggle${qNum}`);
            const otherToggle = document.getElementById(`toggle${otherQ}`);
            const currentSelect = document.getElementById(`q${qNum}_pattern`);
            const otherSelect = document.getElementById(`q${otherQ}_pattern`);

            if (currentToggle.checked) {
                otherToggle.checked = false;
                currentSelect.value = "16";
                if (otherSelect.value === "16") { otherSelect.value = "8+8"; }
            } else {
                currentSelect.value = "8+8";
                otherToggle.checked = true;
                otherSelect.value = "16";
            }
            toggleInputs('7');
            toggleInputs('8');
        }

        function toggleInputs(qNum) {
            const pattern = document.getElementById(`q${qNum}_pattern`).value;
            const toggle = document.getElementById(`toggle${qNum}`);
            const divA = document.getElementById(`div_${qNum}a_ii`);
            const divB = document.getElementById(`div_${qNum}b_ii`);

            if (pattern === "16") {
                toggle.checked = true;
                const otherQ = (qNum === '7') ? '8' : '7';
                document.getElementById(`toggle${otherQ}`).checked = false;
                divA.classList.add('d-none'); divB.classList.add('d-none');
            } else {
                toggle.checked = false;
                divA.classList.remove('d-none'); divB.classList.remove('d-none');
            }
        }

        function syncSub(sourceName, targetName) {
            const source = document.getElementsByName(sourceName)[0];
            const target = document.getElementsByName(targetName)[0];
            if (source && target) { target.value = source.value; }
        }

        function syncPartB() {
            document.getElementById('pb_co_b').value = document.getElementById('pb_co_a').value;
            document.getElementById('pb_k_b').value = document.getElementById('pb_k_a').value;
        }

        window.onload = function() {
            filterCOs(); toggleInputs('7'); toggleInputs('8');
        };
    </script>
</body>
</html>
"""

# --- UPDATED PAPER TEMPLATE (Editable & Split-Row Support) ---
PAPER_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: "Times New Roman", Times, serif; padding: 10px; color: black; line-height: 1.2; }
        [contenteditable="true"]:hover { background-color: #fff9c4; outline: 1px dashed #ccc; }
       
        /* Center the logo container */
        .inst-header {
            text-align: center;
            width: 100%;
            margin-bottom: 15px;
        }
        .inst-logo {
            display: block;
            margin: 0 auto;
            max-width: 450px; /* Increased size */
            height: auto;
        }

        /* Container for Roll No and Reg No below Department */
        .id-container {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin: 15px 0;
            font-weight: bold;
        }
      
        .header-text {
    display: block;          /* Occupies the full line */
    text-align: center;      /* Centers the text within that line */
    width: 100%;             /* Ensures it spans the entire page width */
    font-weight: bold;
    font-size: 13pt;
    margin: 20px 0;
    line-height: 1.6;
    clear: both;             /* Ensures no overlap from the logo above */
}

/* Container for the ID lines */
	 .reg-container {
    display: flex;
    justify-content: space-between; /* Roll No (Left) | Reg No (Right) */
    align-items: flex-end;
    width: 100%;
    margin-top: 10px;
    margin-bottom: 10px;
}

        .reg-table { border-collapse: collapse; margin-bottom: 10px; float: left; }
        .reg-table td { border: 1px solid black; width: 20px; height: 25px; text-align: center; font-weight: bold; }
        .clear { clear: both; }
        .meta-table { width: 100%; border-collapse: collapse; margin-top: 10px; border: 1px solid black; }
        .meta-table td { border: 1px solid black; padding: 5px; font-weight: bold; font-size: 13px; }
        .section-header { font-weight: bold; text-align: center; margin: 15px 0 5px 0; font-size: 14px; }
        .question-table { width: 100%; border-collapse: collapse; margin-top: 5px; }
        .question-table th, .question-table td { border: 1px solid black; padding: 5px; font-size: 13px; vertical-align: top; }
        .bloom-table { width: 100%; border-collapse: collapse; margin-top: 20px; text-align: center; }
        .bloom-table th, .bloom-table td { border: 1px solid black; padding: 4px; font-size: 12px; }
        .footer-sigs { display: flex; justify-content: space-between; margin-top: 50px; font-weight: bold; }
       
        @media print {
            .no-print { display: none; }
            [contenteditable="true"]:hover { background-color: transparent; outline: none; }
        }
    /* This applies to the screen and the printed PDF */
    @media print, screen {
        body, table, td, th, div, p {
            font-family: "Times New Roman", Times, serif !important;
            font-size: 12pt !important;
            color: black;
        }
       
        /* Ensure table borders are visible in PDF */
        table {
            border-collapse: collapse;
            width: 100%;
        }
       
        table, th, td {
            border: 1px solid black !important;
        }

        /* Prevent large headers from overriding size 12 */
        h1, h2, h3, .header-title {
            font-size: 12pt !important;
            font-weight: bold;
            margin: 5px 0;
        }
    }
   
    
    .reg-label { font-weight: bold; margin-right: 10px; }
    .reg-grid {
            display: inline-table;
            border-collapse: collapse;
            margin-left: 10px;
        }
        .reg-grid td {
            border: 1px solid black !important;
            width: 22px;
            height: 25px;
            text-align: center;}

    </style>
</head>
<body>
   

    <div class="inst-header">
        <img src="{{ url_for('static', filename='logo.jpg') }}" class="inst-logo" onerror="this.style.display='none';">
    </div>

       <div class="clear"></div>
   <div class="header-text" contenteditable="true">
        B.E. / B.TECH. DEGREE EXAMINATIONS <br>
        (JANUARY to MAY 2026) <br>
        DEPARTMENT OF ____________________
    </div>

    <div class="reg-container">
        <div style="font-weight:bold; padding-bottom: 5px;">
            Roll No. __________________
        </div>
        
        <div style="display: flex; align-items: center;">
            <span class="reg-label" style="font-weight: bold; margin-right: 10px;">Reg.No.</span>
            <table class="reg-grid" style="border-collapse: collapse;">
                <tr>
                    <td style="border: 1px solid black; width: 22px; height: 28px; text-align: center;">9</td>
                    <td style="border: 1px solid black; width: 22px; height: 28px; text-align: center;">2</td>
                    <td style="border: 1px solid black; width: 22px; height: 28px; text-align: center;">0</td>
                    <td style="border: 1px solid black; width: 22px; height: 28px; text-align: center;">4</td>
                    {% for i in range(8) %}
                    <td style="border: 1px solid black; width: 22px; height: 28px; text-align: center;"></td>
                    {% endfor %}
                </tr>
            </table>
        </div>
    </div>
    <div class="clear"></div>

    

   <table class="meta-table">
        <tr><td colspan="2" contenteditable="true">Internal Assessment: I</td><td colspan="2" contenteditable="true">Semester: Fourth / Sixth </td></tr>
        <tr><td colspan="4" contenteditable="true">Course Code - Course Name: __________________________________ (Common to ____) </td></tr>
        <tr><td contenteditable="true">Regulation: KCET 2021 </td><td contenteditable="true">Max. Marks: 50 Marks </td><td colspan="2" contenteditable="true">Duration: 1 hour 30 minutes </td></tr>
    </table>
    <table class="co-table">
        <thead>
            <tr style="background-color: #f2f2f2;">
                <th style="border: 1px solid black; padding: 8px; width: 15%; text-align: center;">CO Index</th>
                <th style="border: 1px solid black; padding: 8px; text-align: center;">Course Outcomes</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td style="border: 1px solid black; text-align: center; font-weight: bold;">CO1</td>
                <td style="border: 1px solid black; padding: 8px;" contenteditable="true">
                    {{ co_descriptions.get('CO1', '') }}
                </td>
            </tr>
            <tr>
                <td style="border: 1px solid black; text-align: center; font-weight: bold;">CO2</td>
                <td style="border: 1px solid black; padding: 8px;" contenteditable="true">
                    {{ co_descriptions.get('CO2', '') }}
                </td>
            </tr>
        </tbody>
    </table>
   
    <table class="bloom-table">
        <tr style="background:#eee;"><th colspan="5">Marks distribution based on Bloom’s Taxonomy Level </th></tr>
        <tr><td>Remember (K-1)</td><td>Understand (K-2)</td><td>Apply (K-3)</td><td>Analyze (K-4)</td><td>Total</td></tr>
        <tr contenteditable="true">
            <td>{{ k_stats.get('K1', {'marks':0}).marks }}</td>
            <td>{{ k_stats.get('K2', {'marks':0}).marks }}</td>
            <td>{{ k_stats.get('K3', {'marks':0}).marks }}</td>
            <td>{{ k_stats.get('K4', {'marks':0}).marks }}</td>
            <td>90 </td>
        </tr>
    </table>

    <div class="section-header">Answer all the Questions </div>

    <table class="question-table">
    <tr style="background:#eee;">
        <th width="12%">CO, BTL</th> <th width="8%">Q. No.</th>
        <th width="68%">Part A (5 x 2 = 10 Marks)</th>
        <th width="12%">Marks</th>
    </tr>
    {% for q in part_a %}
    <tr>
        <td align="center">{{ q['CO'] }}, {{ q['BTL'] }}</td>
       
        <td align="center">{{ loop.index }}</td>
        <td contenteditable="true">
            <div>{{ q['Question_Text'] }}</div>
            {% if loop.index == 2 or loop.index == 4 %}
                {% endif %}
        </td>
        <td align="center">2</td>
    </tr>
    {% endfor %}
        <tr style="background:#eee;"><th colspan="4">Part B (1 x 8 = 8 Marks) </th></tr>
        <tr>
            <td align="center" contenteditable="true">{{ p_b[0]['CO'] }}, {{ p_b[0]['BTL'] }}</td>
            <td align="center" contenteditable="true">6 (a)</td>
            <td contenteditable="true">{{ p_b[0]['Question'] }}</td>
            <td align="center" contenteditable="true">8</td>
        </tr>
        <tr><td colspan="4" align="center"><b>(OR) </b></td></tr>
        <tr>
            <td align="center" contenteditable="true">{{ p_b[1]['CO'] }}, {{ p_b[1]['BTL'] }}</td>
            <td align="center" contenteditable="true">6 (b)</td>
            <td contenteditable="true">{{ p_b[1]['Question'] }}</td>
            <td align="center" contenteditable="true">8</td>
        </tr>

        <tr style="background:#eee;"><th colspan="4">Part C (2 x 16 = 32 Marks) </th></tr>
        {% for qnum, data in [('7', p_c7), ('8', p_c8)] %}
            {# --- Option (a) Split rows --- #}
            {% for sq in data['a'] %}
            <tr>
                <td align="center" contenteditable="true">{{ sq['CO'] }}, {{ sq['BTL'] }}</td>
                <td align="center" contenteditable="true">{{ qnum }} (a) {% if data['a']|length > 1 %}({{ loop.index|lower }}){% endif %}</td>
                <td contenteditable="true">{{ sq['Question'] }}</td>
                <td align="center" contenteditable="true">{{ sq['Marks'] }}</td>
            </tr>
            {% endfor %}
           
            <tr><td colspan="4" align="center"><b>(OR) </b></td></tr>
           
            {# --- Option (b) Split rows --- #}
            {% for sq in data['b'] %}
            <tr>
                <td align="center" contenteditable="true">{{ sq['CO'] }}, {{ sq['BTL'] }}</td>
                <td align="center" contenteditable="true">{{ qnum }} (b) {% if data['b']|length > 1 %}({{ loop.index|lower }}){% endif %}</td>
                <td contenteditable="true">{{ sq['Question'] }}</td>
                <td align="center" contenteditable="true">{{ sq['Marks'] }}</td>
            </tr>
            {% endfor %}
        {% endfor %}
    </table>

    <div class="no-print" style="text-align:center; padding: 20px; background: #f8f9fa; margin-bottom: 20px; border: 2px dashed #007bff; border-radius: 8px;">
    <h4 class="text-primary">Download as Word (.docx)</h4>
    <p>Upload your college's official Word template to populate it with this data.</p>
   
    <form action="/download/word" method="post" enctype="multipart/form-data" style="display: inline-block;">
        <input type="hidden" name="paper_json" id="paper_json">
       
        <div style="display: flex; gap: 10px; align-items: center; justify-content: center;">
            <input type="file" name="template_file" class="form-control" accept=".docx" required style="width: 300px;">
            <button type="submit" onclick="prepareData()" class="btn btn-success">
                Generate Word File
            </button>
            <button type="button" onclick="window.print()" class="btn btn-secondary">
                Print to PDF
            </button>
        </div>
    </form>
</div>

<script>
    function prepareData() {
        // Collects the data from the current page (including any manual edits made)
        const data = {
            part_a: {{ part_a|tojson }},
            p_b: {{ p_b|tojson }},
            p_c7: {{ p_c7|tojson }},
            p_c8: {{ p_c8|tojson }},
            k_stats: {{ k_stats|tojson }}
        };
        document.getElementById('paper_json').value = JSON.stringify(data);
    }
</script>

<script>
    function handleDownload() {
        const format = document.getElementById('downloadFormat').value;
        if (format === 'word') {
            window.location.href = '/download/word';
        } else {
            // For PDF, we just use the browser's print-to-pdf feature
            window.print();
        }
    }
</script>
</body>
</html>
"""

def process_pdf(file_storage):
    extracted = []
    with pdfplumber.open(file_storage) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                # Clean the row
                clean_row = [str(c) if c else "" for c in row]
                blob = " ".join(clean_row)
               
                # Metadata checks
                co = re.search(r'(CO\d)', blob, re.I)
                btl = re.search(r'(K\d)', blob, re.I)
                marks = re.search(r'\b(2|6|8|10|16)\b', blob)
               
                if co and btl and marks:
                    q_raw = max(clean_row, key=len).strip()
                   
                    # --- RE-INDENT THIS LINE CAREFULLY WITH SPACES ---
                    q_full = " ".join(q_raw.split())
                   
                    extracted.append({
                        'Question': q_full,
                        'CO': co.group(1).upper(),
                        'BTL': btl.group(1).upper(),
                        'Marks': int(marks.group(1))
                    })
    return pd.DataFrame(extracted)
TEMP_DF = None

@app.route('/')
def index():
    return render_template_string(UPLOAD_HTML)

@app.route('/setup', methods=['POST'])
def setup_bank():
    global TEMP_DF
    file = request.files['file']
    try:
        # 1. Load the Question Bank File
        if file.filename.lower().endswith('.pdf'):
            df = process_pdf(file)
        elif file.filename.lower().endswith('.xlsx'):
            df = pd.read_excel(file, dtype=str)
        else:
            df = pd.read_csv(file, encoding='latin1', dtype=str)

        # 2. Standardize column names to uppercase
        df.columns = df.columns.astype(str).str.strip().str.upper()

        # 3. Extract CO Statements (CO1-CO5) for the Outcome Table
        co_map = {}
        for index, row in df.iterrows():
            # Standardize the value to catch variations like " co1 " or "CO1"
            first_col_val = str(row.iloc[0]).strip().upper().replace(" ", "")
            if re.match(r'^CO[1-5]$', first_col_val):
                description = str(row.iloc[1]).strip()
                if description:
                    co_map[first_col_val] = description
       
        session['co_descriptions'] = co_map
       

        # 4. Standardize Columns for Question Filtering
        rename_map = {'MARK': 'Marks', 'QUESTION': 'Question', 'CO': 'CO', 'BTL': 'BTL'}
        df = df.rename(columns=lambda x: next((v for k, v in rename_map.items() if k in x), x))

        # 5. Clean and format the data
        df['CO'] = df['CO'].astype(str).str.replace(' ', '').str.upper()
        df['BTL'] = df['BTL'].astype(str).str.replace(' ', '').str.upper()
        df['Marks'] = pd.to_numeric(df['Marks'], errors='coerce')
       
        # Determine Unit number based on the CO string (e.g., CO1 -> Unit 1)
        df['Unit'] = df['CO'].apply(lambda x: re.search(r'\d+', x).group() if re.search(r'\d+', x) else "1")
       
        # 6. Finalize TEMP_DF and prepare UI Mapping
        TEMP_DF = df.dropna(subset=['Question', 'Marks'])
        unit_to_cos = {str(i): sorted(TEMP_DF[TEMP_DF['Unit'] == str(i)]['CO'].unique().tolist()) for i in range(1, 6)}
       
        return render_template_string(CONFIG_HTML, mapping=unit_to_cos, k_levels=sorted(TEMP_DF['BTL'].unique().tolist()))
   
    except Exception as e:
        return f"Setup Error: {str(e)}"

@app.route('/generate', methods=['POST'])
def generate_paper():
    global TEMP_DF
    if TEMP_DF is None: return redirect('/')
   
    used_indices = set()
    k_stats = {}

    def pick_q(co, k, marks, count, force_mcq=False):
        nonlocal used_indices, k_stats
        # STRICT FILTERING: CO, Marks, and K-Level must all match
        filters = (TEMP_DF['CO'] == co) & (TEMP_DF['Marks'] == marks) & (TEMP_DF['BTL'] == k)
        is_mcq_content = TEMP_DF['Question'].str.contains(r'\b[Aa]\)', na=False)
        pool = TEMP_DF[filters & is_mcq_content] if force_mcq else TEMP_DF[filters & ~is_mcq_content]
        pool = pool[~pool.index.isin(used_indices)]
       
        if pool.empty: return []
        res = pool.sample(n=min(count, len(pool)))
        data_list = res.to_dict('records')

        for r in data_list:
            used_indices.add(res.index[0])
            btl_val = r.get('BTL', 'K1')
            if btl_val not in k_stats: k_stats[btl_val] = {'marks': 0}
            k_stats[btl_val]['marks'] += marks
            clean_q = " ".join(r['Question'].split())
            if re.search(r'\b[Aa]\)', clean_q):
                stem_parts = re.split(r'\s*\b[Aa]\)', clean_q, maxsplit=1)
                r['Question_Text'] = stem_parts[0].strip()
                options_found = re.findall(r'[A-Da-d]\)\s*(.*?)(?=[A-Da-d]\)|Justify|$)', clean_q)
                r['options'] = [opt.strip() for opt in options_found if opt.strip()][:4]
            else:
                r['Question_Text'] = clean_q
                r['options'] = ["", "", "", ""]
        return data_list

    try:
        # PART A
        p_a = []
        for i in range(1, 6):
            co_val = request.form.get(f'pa_co{i}')
            k_val = request.form.get(f'pa_k{i}')
            q_list = pick_q(co_val, k_val, 2, 1, force_mcq=(i == 2 or i == 4))
            p_a.append(q_list[0] if q_list else {
                'Question_Text': f"⚠️ Not enough questions in {co_val} - {k_val}",
                'CO': '-', 'BTL': '-', 'options': ["","","",""]
            })

        # PART B
        co_a, k_a = request.form.get('pb_co_a'), request.form.get('pb_k_a')
        co_b, k_b = request.form.get('pb_co_b'), request.form.get('pb_k_b')
        list_6a = pick_q(co_a, k_a, 8, 1)
        list_6b = pick_q(co_b, k_b, 8, 1)
        p_b = [
            list_6a[0] if list_6a else {'Question': f"⚠️ Not enough questions in {co_a} - {k_a}", 'CO': '-', 'BTL': '-'},
            list_6b[0] if list_6b else {'Question': f"⚠️ Not enough questions in {co_b} - {k_b}", 'CO': '-', 'BTL': '-'}
        ]

        # PART C
        def get_pc(q_num):
            pat = request.form.get(f'q{q_num}_pattern')
            def fetch(pref):
                co = request.form.get(f'pc_co{q_num}{pref}i')
                k = request.form.get(f'pc_k{q_num}{pref}i')
                if pat == "16":
                    q = pick_q(co, k, 16, 1)
                    return q if q else [{'Question': f"⚠️ Not enough questions in {co} - {k}", 'CO': '-', 'BTL': '-', 'Marks': 16}]
               
                m_list = [8, 8] if pat == "8+8" else [10, 6]
                co_ii = request.form.get(f'pc_co{q_num}{pref}ii')
                k_ii = request.form.get(f'pc_k{q_num}{pref}ii')
                q1 = pick_q(co, k, m_list[0], 1)
                q2 = pick_q(co_ii, k_ii, m_list[1], 1)
                return [
                    q1[0] if q1 else {'Question': f"⚠️ Not enough questions in {co} - {k}", 'CO': '-', 'BTL': '-', 'Marks': m_list[0]},
                    q2[0] if q2 else {'Question': f"⚠️ Not enough questions in {co_ii} - {k_ii}", 'CO': '-', 'BTL': '-', 'Marks': m_list[1]}
                ]
            return {"a": fetch("a"), "b": fetch("b")}

        paper_results = {
            'part_a': p_a,
            'p_b': p_b,
            'p_c7': get_pc("7"),
            'p_c8': get_pc("8"),
            'k_stats': dict(sorted(k_stats.items()))
        }
        # ... (after paper_results dictionary is created)
        session['last_paper'] = paper_results
       
        # Get descriptions from session
        co_descriptions = session.get('co_descriptions', {})
       
        # FIX: Explicitly pass co_descriptions to the template
        return render_template_string(PAPER_TEMPLATE, co_descriptions=co_descriptions, **paper_results)
    except Exception as e:
        return f"<h1>Logic Error</h1><p>{str(e)}</p>"

@app.route('/download/word', methods=['POST'])
def download_word_file():
    raw_data = request.form.get('paper_json')
    if not raw_data: return "No data", 400
    paper_data = json.loads(raw_data)
    template_file = request.files.get('template_file')
    if not template_file: return "Upload template", 400
   
    try:
        doc = Document(template_file)

        # Main Replacement Function
        def fill(tag, val):
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if tag in cell.text:
                            # Standard text replacement within cells
                            cell.text = cell.text.replace(tag, str(val))

       
        # --- FILL COURSE OUTCOMES (Manual 2 Rows) ---
        co_data = session.get('co_descriptions', {})
       
        # Manually fill only CO1 and CO2
        fill("{{CO1_DESC}}", co_data.get('CO1', ""))
        fill("{{CO2_DESC}}", co_data.get('CO2', ""))

        # --- 1. FILL BLOOM'S TAXONOMY MARKS ---
        stats = paper_data.get('k_stats', {})
        total_marks = 0
        for k_level in ['K1', 'K2', 'K3', 'K4']:
            # Safe access to nested dictionary marks
            k_data = stats.get(k_level, {})
            # Ensure we are handling a dictionary from the JSON [cite: 6]
            m = k_data.get('marks', 0) if isinstance(k_data, dict) else 0
            fill(f"{{{{{k_level}_M}}}}", m)
            total_marks += m
        fill("{{TOTAL_M}}", total_marks)

        # --- 2. FILL PART A (1-5) ---
        for i, q in enumerate(paper_data['part_a']):
            n = i + 1
            text = q.get('Question_Text', q.get('Question', ''))
            # Format MCQs specifically for Part A [cite: 6]
            if n in [2, 4] and q.get('options'):
                opts = q['options']
                text += f"\na) {opts[0]} b) {opts[1]} c) {opts[2]} d) {opts[3]}\nJustify your answer."
            fill(f"{{{{Q{n}_T}}}}", text)
            fill(f"{{{{Q{n}_C}}}}", f"{q.get('CO', '-')}, {q.get('BTL', '-')}")

        # --- 3. FILL PART B (6a, 6b) ---
        p_b = paper_data.get('p_b', [])
        for idx, sub in enumerate(['A', 'B']):
            if idx < len(p_b):
                fill(f"{{{{Q6{sub}_T}}}}", p_b[idx].get('Question', ''))
                fill(f"{{{{Q6{sub}_C}}}}", f"{p_b[idx].get('CO', '-')}, {p_b[idx].get('BTL', '-')}")

        # --- 4. FILL PART C (7 & 8 with choice tags) ---
        def fill_part_c(q_data, q_num):
            for choice in ['a', 'b']:
                subs = q_data.get(choice, [])
                # Populates main choice CO/BTL (e.g., {{Q7A_C}}) [cite: 6]
                if subs:
                    main_c = f"{subs[0].get('CO', '-')}, {subs[0].get('BTL', '-')}"
                    fill(f"{{{{Q{q_num}{choice.upper()}_C}}}}", main_c)

                # Populates subdivisions (i and ii)
                for i, sub in enumerate(subs):
                    suffix = "i" if i == 0 else "ii"
                    tag_prefix = f"Q{q_num}{choice.upper()}{suffix}"
                    fill(f"{{{{{tag_prefix}_T}}}}", sub.get('Question', ''))
                    fill(f"{{{{{tag_prefix}_C}}}}", f"{sub.get('CO', '-')}, {sub.get('BTL', '-')}")

        fill_part_c(paper_data.get('p_c7', {}), 7)
        fill_part_c(paper_data.get('p_c8', {}), 8)

        # --- 5. CLEANUP ---
        fill("{{DEPT}}", ""); fill("{{SEM}}", ""); fill("{{COURSE_INFO}}", "")

        target_file = io.BytesIO()
        doc.save(target_file)
        target_file.seek(0)
        return send_file(target_file, as_attachment=True, download_name="Official_Paper.docx")

    except Exception as e:
        return f"Download Error: {str(e)}"
if __name__ == '__main__':
    app.run(debug=True)
