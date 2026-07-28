# app.py
import os
import uuid
import calendar as cal
from datetime import datetime
import pandas as pd
import pdfplumber
from pathlib import Path
from flask import Flask, request, jsonify, render_template_string
from werkzeug.utils import secure_filename
import re

app = Flask(__name__)
UPLOAD_FOLDER = Path("temp_uploads")
UPLOAD_FOLDER.mkdir(exist_ok=True)

# ─── MONTH DETECTION ────────────────────────────────────────

MONTH_NAMES = {
    "jan": 1, "january": 1, "januari": 1,
    "feb": 2, "february": 2, "februari": 2,
    "mar": 3, "march": 3, "maret": 3,
    "apr": 4, "april": 4,
    "may": 5, "mei": 5,
    "jun": 6, "june": 6, "juni": 6,
    "jul": 7, "july": 7, "juli": 7,
    "aug": 8, "august": 8, "agustus": 8, "agu": 8, "ags": 8,
    "sep": 9, "september": 9, "sept": 9,
    "oct": 10, "october": 10, "okt": 10, "oktober": 10,
    "nov": 11, "november": 11, "nop": 11, "nopember": 11,
    "dec": 12, "december": 12, "des": 12, "desember": 12,
}


def detect_months_in_text(text):
    text = str(text).lower().strip()
    found = set()

    if not text or text in ("nan", "none", ""):
        return found

    # Skip if text is just a plain number (row numbers, quantities, etc.)
    clean = text.replace(",", "").replace(".", "").strip()
    if clean.isdigit():
        return found

    # Pattern 1: "Jan-26", "Dec 2025", "Januari-26", "Feb/26"
    for match in re.finditer(r"([a-z]{3,})[\s\-_/\.]+(\d{2,4})", text):
        name = match.group(1)
        yr = int(match.group(2))
        if yr < 100:
            yr += 2000
        mn = MONTH_NAMES.get(name)
        if mn and 2020 <= yr <= 2035:
            found.add((yr, mn))

    # Pattern 2: "2025-01", "2025/12"
    for match in re.finditer(r"(\d{4})[\s\-_/\.]+(\d{1,2})(?!\d)", text):
        yr = int(match.group(1))
        mn = int(match.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035:
            found.add((yr, mn))

    # Pattern 3: "2025 January", "2025-Jan"
    for match in re.finditer(r"(\d{4})[\s\-_/\.]+([a-z]{3,})", text):
        yr = int(match.group(1))
        mn = MONTH_NAMES.get(match.group(2))
        if mn and 2020 <= yr <= 2035:
            found.add((yr, mn))

    # Pattern 4: Full dates with 4-digit year only
    # "01/05/2026", "2026-05-01"
    for match in re.finditer(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", text):
        a, b, c = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if 2020 <= c <= 2035:
            if 1 <= b <= 12:
                found.add((c, b))
            if 1 <= a <= 12 and a != b:
                found.add((c, a))

    for match in re.finditer(r"(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})", text):
        yr, mn = int(match.group(1)), int(match.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035:
            found.add((yr, mn))

    return found


def detect_months_in_excel(rows):
    found = set()
    for row in rows:
        for cell in row:
            s = str(cell).strip()
            if s and s not in ("", "nan", "None"):
                found.update(detect_months_in_text(s))
    return found


def detect_months_in_lines(lines):
    found = set()
    for line in lines:
        found.update(detect_months_in_text(line))
    return found


def validate_month(detected, target_year, target_month):
    target = (target_year, target_month)
    target_label = cal.month_name[target_month] + " " + str(target_year)

    if not detected:
        return {
            "status": "warning", "icon": "⚠️",
            "message": "No month/year detected. Cannot verify.",
            "match": False, "detected": [], "target": target_label,
        }

    detected_labels = sorted([
        cal.month_name[m] + " " + str(y) for y, m in detected
    ])

    if target in detected:
        if len(detected) == 1:
            return {
                "status": "ok", "icon": "✅",
                "message": "Matches: " + target_label,
                "match": True, "detected": detected_labels, "target": target_label,
            }
        else:
            others = [l for l in detected_labels if l != target_label]
            return {
                "status": "ok_multi", "icon": "⚠️",
                "message": "Contains " + target_label + " but also has: " + ", ".join(others),
                "match": True, "detected": detected_labels, "target": target_label,
            }
    else:
        return {
            "status": "mismatch", "icon": "⚠️",
            "message": "WRONG MONTH — Expected " + target_label + " but found: " + ", ".join(detected_labels),
            "match": False, "detected": detected_labels, "target": target_label,
        }


# ─── FILE READERS ────────────────────────────────────────────

def read_excel(path):
    try:
        engine = "xlrd" if path.suffix == ".xls" else "openpyxl"
        df = pd.read_excel(path, header=None, engine=engine).fillna("")
        rows = []
        for row in df.values.tolist():
            rows.append([str(c) if str(c) != "" else "" for c in row])
        return {"type": "table", "rows": rows, "cols": len(rows[0]) if rows else 0}
    except Exception as e:
        return {"type": "error", "message": "Excel Error: " + str(e)}


def smart_ocr_to_table(lines):
    """
    Convert OCR text lines into a structured table.

    Expected pattern:
      1. 01 March 2026, Sunday 330.685.175 Z 330.247.275

    Output columns:
      No | Date | Day | Value 1 | Value 2 | Noise
    """
    import re

    rows = []

    # Keep useful metadata lines at the top
    meta_lines = []

    # Data row pattern:
    # 1. 01 March 2026, Sunday 330.685.175 Z 330.247.275
    row_pattern = re.compile(
        r"^\s*(\d+)[\.,]\s+"                       # 1.
        r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})"          # 01 March 2026
        r",?\s+([A-Za-z]+)\s+"                    # Sunday
        r"(.*)$"                                  # the rest
    )

    # Money-like number pattern:
    money_pattern = re.compile(r"(?<!\d)(\d{1,3}(?:[.,]\d{3})+|\d{6,})(?!\d)")

    for line in lines:
        clean = str(line).strip()
        if not clean:
            continue

        m = row_pattern.match(clean)

        if m:
            no = m.group(1)
            full_date = m.group(2)
            day_name = m.group(3)
            tail = m.group(4)

            # Find all large numeric values in tail
            amounts = money_pattern.findall(tail)

            value_1 = amounts[0] if len(amounts) > 0 else ""
            value_2 = amounts[1] if len(amounts) > 1 else ""

            # Remove extracted amounts from tail to leave OCR noise/markers
            noise = tail
            for amt in amounts[:2]:
                noise = noise.replace(amt, " ", 1)

            noise = re.sub(r"\s+", " ", noise).strip(" -_:;,.")
            rows.append([no, full_date, day_name, value_1, value_2, noise])

        else:
            # Treat non-data lines as metadata
            meta_lines.append(clean)

    # If no structured rows found, return None so caller can fall back
    if not rows:
        return None, 0

    # Put metadata above the table as padded rows
    final_rows = []

    for meta in meta_lines[:5]:
        final_rows.append([meta, "", "", "", "", ""])

    # Optional header row inside preview
    final_rows.append(["No", "Date", "Day", "Value 1", "Value 2", "Noise"])
    final_rows.extend(rows)

    return final_rows, 6


def read_pdf_ocr(path):
    """Fallback: Convert PDF to images and run Tesseract OCR."""
    try:
        from pdf2image import convert_from_path
        import pytesseract

        images = convert_from_path(path, dpi=300)
        text_lines = []

        for img in images:
            text = pytesseract.image_to_string(img, lang="eng")
            if text:
                text_lines.extend([l.strip() for l in text.split("\n") if l.strip()])

        if text_lines:
            # ─── NEW: Convert raw text lines to a table ───────────
            rows, cols = smart_ocr_to_table(text_lines)
            if rows:
                return {
                    "type": "table", 
                    "rows": rows, 
                    "cols": cols,
                    "is_ocr": True # Flag to show it was OCR'd
                }
            return {"type": "text", "lines": text_lines}

        return {"type": "error", "message": "PDF OCR: No text detected."}
    except Exception as e:
        return {"type": "error", "message": "PDF OCR Error: " + str(e)}


def read_pdf(path):
    try:
        tables_found = []
        text_lines = []

        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    for table in page_tables:
                        for row in table:
                            cleaned = [str(c).strip() if c else "" for c in row]
                            # Only add if row isn't mostly empty
                            if any(cleaned):
                                tables_found.append(cleaned)
                
                # Also capture text for month detection
                text = page.extract_text()
                if text:
                    text_lines.extend([l.strip() for l in text.split("\n") if l.strip()])

        # Case 1: Digital PDF with structured tables
        if tables_found:
            max_cols = max(len(r) for r in tables_found)
            for r in tables_found:
                while len(r) < max_cols:
                    r.append("")
            return {"type": "table", "rows": tables_found, "cols": max_cols}

        # Case 2: Digital PDF with only text (no table structure detected)
        if text_lines and len(text_lines) > 5: # Some threshold to skip empty pages
             # Try to see if the text lines look like a table
             rows, cols = smart_ocr_to_table(text_lines)
             if rows:
                 return {"type": "table", "rows": rows, "cols": cols}
             return {"type": "text", "lines": text_lines}

        # Case 3: Scanned PDF → Trigger OCR
        return read_pdf_ocr(path)

    except Exception as e:
        return read_pdf_ocr(path)

# ─── HTML ────────────────────────────────────────────────────

HTML_UI = r"""
<!DOCTYPE html>
<html>
<head>
    <title>Tenant Report Reader</title>
    <style>
        body {
            font-family: sans-serif;
            background: #0f172a;
            color: #e2e8f0;
            padding: 40px;
            margin: 0;
        }
        .container { max-width: 1200px; margin: 0 auto; }
        h1 { margin-bottom: 5px; }
        .subtitle { color: #94a3b8; margin-bottom: 25px; }

        /* Config */
        .config-card {
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .config-title {
            font-size: 1rem;
            font-weight: 600;
            color: #93c5fd;
            margin-bottom: 15px;
        }
        .config-row {
            display: flex;
            gap: 15px;
            align-items: flex-end;
            flex-wrap: wrap;
        }
        .config-group label {
            display: block;
            font-size: 0.8rem;
            color: #94a3b8;
            margin-bottom: 5px;
        }
        .config-group select,
        .config-group input[type="number"] {
            background: #263248;
            border: 1px solid #475569;
            border-radius: 8px;
            color: #e2e8f0;
            padding: 8px 12px;
            font-size: 0.95rem;
            outline: none;
        }
        .config-group select option { background: #1e293b; }
        .month-preview {
            font-size: 0.85rem;
            color: #94a3b8;
            padding: 8px 0;
        }
        .month-preview b { color: #3b82f6; }

        /* Drop Zone */
        .drop-zone {
            border: 3px dashed #3b82f6;
            border-radius: 20px;
            padding: 50px;
            text-align: center;
            cursor: pointer;
            background: rgba(59, 130, 246, 0.05);
            transition: 0.3s;
        }
        .drop-zone.hover {
            background: rgba(59, 130, 246, 0.15);
            border-color: #60a5fa;
            transform: scale(1.01);
        }
        .drop-zone .icon { font-size: 3rem; margin-bottom: 10px; }
        .drop-zone p { color: #94a3b8; margin: 5px 0; }
        .drop-zone b { color: #93c5fd; }

        /* Browse Buttons */
        .browse-row {
            display: flex;
            gap: 10px;
            justify-content: center;
            margin-top: 15px;
            flex-wrap: wrap;
        }
        .browse-btn {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 10px 20px;
            border-radius: 10px;
            font-size: 0.9rem;
            font-weight: 600;
            border: none;
            cursor: pointer;
            transition: all 0.2s;
        }
        .browse-btn:hover { transform: translateY(-1px); }
        .btn-files {
            background: #3b82f6;
            color: #fff;
        }
        .btn-files:hover { background: #2563eb; }
        .btn-folder {
            background: #10b981;
            color: #fff;
        }
        .btn-folder:hover { background: #059669; }
        .btn-clear {
            background: transparent;
            color: #94a3b8;
            border: 1px solid #475569;
        }
        .btn-clear:hover { background: #263248; }

        .hidden-input { display: none; }

        /* Loader */
        .loader {
            display: none;
            text-align: center;
            color: #3b82f6;
            font-weight: bold;
            margin-top: 15px;
            padding: 15px;
        }
        .spinner {
            display: inline-block;
            width: 14px; height: 14px;
            border: 2px solid #3b82f6;
            border-top: 2px solid transparent;
            border-radius: 50%;
            animation: spin 0.6s linear infinite;
            margin-right: 8px;
            vertical-align: middle;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        /* Summary */
        .summary { margin-top: 20px; display: none; }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 10px;
        }
        .summary-card {
            background: #1e293b;
            border: 1px solid #334155;
            padding: 15px;
            border-radius: 10px;
            text-align: center;
        }
        .summary-card .num { font-size: 1.5rem; font-weight: bold; }
        .summary-card .label { font-size: 0.75rem; color: #94a3b8; margin-top: 4px; }
        .c-blue { color: #3b82f6; }
        .c-green { color: #10b981; }
        .c-red { color: #ef4444; }
        .c-yellow { color: #f59e0b; }
        .c-purple { color: #a78bfa; }

        /* File Cards */
        .file-list { margin-top: 20px; }
        .file-card {
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 12px;
            margin-bottom: 12px;
            overflow: hidden;
        }
        .file-card.mismatch { border-color: #ef4444; }
        .file-card.matched { border-color: #10b981; }
        .file-card.warning { border-color: #f59e0b; }

        .file-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 12px 15px;
            cursor: pointer;
            transition: background 0.2s;
        }
        .file-header:hover { background: #263248; }
        .file-info { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }

        .badge {
            font-size: 0.7em;
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: bold;
            letter-spacing: 0.5px;
        }
        .b-xlsx { background: #2563eb; color: #fff; }
        .b-xls  { background: #7c3aed; color: #fff; }
        .b-pdf  { background: #f59e0b; color: #000; }
        .b-ok   { background: rgba(16,185,129,0.2); color: #10b981; border: 1px solid #10b981; }
        .b-fail { background: rgba(239,68,68,0.2); color: #ef4444; border: 1px solid #ef4444; }

        .month-badge {
            font-size: 0.72em;
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: bold;
        }
        .mb-match { background: rgba(16,185,129,0.15); color: #10b981; border: 1px solid rgba(16,185,129,0.3); }
        .mb-mismatch { background: rgba(239,68,68,0.15); color: #ef4444; border: 1px solid rgba(239,68,68,0.3); }
        .mb-warn { background: rgba(245,158,11,0.15); color: #f59e0b; border: 1px solid rgba(245,158,11,0.3); }

        .row-count { font-size: 0.82em; color: #94a3b8; }
        .arrow { color: #64748b; transition: transform 0.2s; }
        .arrow.open { transform: rotate(180deg); }

        /* Month bar */
        .month-bar {
            padding: 8px 15px;
            font-size: 0.82em;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .month-bar.match { background: rgba(16,185,129,0.08); color: #6ee7b7; border-top: 1px solid rgba(16,185,129,0.15); }
        .month-bar.mismatch { background: rgba(239,68,68,0.08); color: #fca5a5; border-top: 1px solid rgba(239,68,68,0.15); }
        .month-bar.warn { background: rgba(245,158,11,0.08); color: #fde68a; border-top: 1px solid rgba(245,158,11,0.15); }
        .month-bar .detected-list {
            font-size: 0.9em;
            color: #94a3b8;
            margin-left: auto;
        }

        /* Preview */
        .preview-area {
            display: none;
            border-top: 1px solid #334155;
            max-height: 500px;
            overflow: auto;
            background: #0c1222;
        }
        .data-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.8em;
            font-family: "Consolas", "Courier New", monospace;
        }
        .data-table th {
            background: #1a3a5c;
            color: #93c5fd;
            padding: 6px 10px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 1;
            border-bottom: 2px solid #2563eb;
            white-space: nowrap;
        }
        .data-table td {
            padding: 5px 10px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            color: #cbd5e1;
            white-space: nowrap;
        }
        .data-table tr:hover td { background: rgba(59, 130, 246, 0.08); }
        .data-table .row-num {
            color: #475569;
            text-align: right;
            padding-right: 12px;
            font-size: 0.85em;
            user-select: none;
            border-right: 1px solid #334155;
            background: #111827;
        }
        .data-table .cell-empty { color: #334155; font-style: italic; }

        .text-preview {
            padding: 15px;
            font-family: monospace;
            font-size: 0.8em;
            color: #10b981;
            white-space: pre-wrap;
            line-height: 1.6;
        }
        .text-preview .line-num {
            display: inline-block;
            width: 40px;
            color: #475569;
            text-align: right;
            margin-right: 12px;
            user-select: none;
        }
        .error-preview { padding: 15px; color: #ef4444; font-size: 0.85em; }
    </style>
</head>
<body>
<div class="container">
    <h1>📂 Tenant Report Reader</h1>
    <p class="subtitle">Upload tenant reports — drop a folder, select files, or both.</p>

    <!-- Month Selector -->
    <div class="config-card">
        <div class="config-title">📅 Report Month</div>
        <div class="config-row">
            <div class="config-group">
                <label>Month</label>
                <select id="monthSel">
                    <option value="1">January</option>
                    <option value="2">February</option>
                    <option value="3">March</option>
                    <option value="4">April</option>
                    <option value="5">May</option>
                    <option value="6">June</option>
                    <option value="7">July</option>
                    <option value="8">August</option>
                    <option value="9">September</option>
                    <option value="10">October</option>
                    <option value="11">November</option>
                    <option value="12">December</option>
                </select>
            </div>
            <div class="config-group">
                <label>Year</label>
                <input id="yearIn" type="number" value="2026" min="2000" max="2100" style="width:90px"/>
            </div>
            <div class="month-preview">
                Target: <b id="targetLabel">January 2026</b>
            </div>
        </div>
    </div>

    <!-- Upload Area -->
    <div class="drop-zone" id="dropZone">
        <div class="icon">📥</div>
        <p><b>Drop files or folder here</b></p>
        <p style="font-size: 0.85em;">Or use the buttons below</p>
    </div>

    <div class="browse-row">
        <button class="browse-btn btn-files" onclick="document.getElementById('fileInput').click()">
            📄 Browse Files
        </button>
        <button class="browse-btn btn-folder" onclick="document.getElementById('folderInput').click()">
            📁 Browse Folder
        </button>
        <button class="browse-btn btn-clear" onclick="clearAll()">
            ↺ Clear All
        </button>
    </div>

    <input type="file" id="fileInput" class="hidden-input" accept=".xlsx,.xls,.pdf" multiple />
    <input type="file" id="folderInput" class="hidden-input" webkitdirectory />

    <div id="loader" class="loader">
        <span class="spinner"></span> Reading and validating files...
    </div>

    <!-- Summary -->
    <div class="summary" id="summary">
        <div class="summary-grid">
            <div class="summary-card">
                <div class="num c-blue" id="sTotal">0</div>
                <div class="label">Total Files</div>
            </div>
            <div class="summary-card">
                <div class="num c-green" id="sOk">0</div>
                <div class="label">Read OK</div>
            </div>
            <div class="summary-card">
                <div class="num c-red" id="sFail">0</div>
                <div class="label">Read Failed</div>
            </div>
            <div class="summary-card">
                <div class="num c-purple" id="sMatch">0</div>
                <div class="label">Month Match ✅</div>
            </div>
            <div class="summary-card">
                <div class="num c-yellow" id="sWrong">0</div>
                <div class="label">Wrong Month ❌</div>
            </div>
        </div>
    </div>

    <div class="file-list" id="fileList"></div>
</div>

<script>
var dropZone    = document.getElementById("dropZone");
var fileList    = document.getElementById("fileList");
var loader      = document.getElementById("loader");
var summaryEl   = document.getElementById("summary");
var fileInput   = document.getElementById("fileInput");
var folderInput = document.getElementById("folderInput");

// Month label
function updateLabel() {
    var months = ["","January","February","March","April","May","June",
                  "July","August","September","October","November","December"];
    var m = parseInt(document.getElementById("monthSel").value);
    var y = document.getElementById("yearIn").value;
    document.getElementById("targetLabel").textContent = months[m] + " " + y;
}
document.getElementById("monthSel").addEventListener("change", updateLabel);
document.getElementById("yearIn").addEventListener("input", updateLabel);
var now = new Date();
document.getElementById("monthSel").value = now.getMonth() + 1;
document.getElementById("yearIn").value = now.getFullYear();
updateLabel();

// ─── DRAG AND DROP (files or folders) ───────────────────────

dropZone.addEventListener("dragover", function(e) {
    e.preventDefault();
    dropZone.classList.add("hover");
});
dropZone.addEventListener("dragleave", function() {
    dropZone.classList.remove("hover");
});

dropZone.addEventListener("drop", async function(e) {
    e.preventDefault();
    dropZone.classList.remove("hover");

    var items = e.dataTransfer.items;
    var formData = new FormData();

    for (var i = 0; i < items.length; i++) {
        var entry = items[i].webkitGetAsEntry();
        if (entry) {
            await walkTree(entry, formData, "");
        }
    }

    sendFiles(formData);
});

// ─── BROWSE FILES BUTTON ────────────────────────────────────

fileInput.addEventListener("change", function() {
    var files = fileInput.files;
    if (!files.length) return;

    var formData = new FormData();
    for (var i = 0; i < files.length; i++) {
        var f = files[i];
        var name = f.name.toLowerCase();
        if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".pdf")) {
            formData.append("files", f, f.name);
        }
    }
    fileInput.value = "";
    sendFiles(formData);
});

// ─── BROWSE FOLDER BUTTON ───────────────────────────────────

folderInput.addEventListener("change", function() {
    var files = folderInput.files;
    if (!files.length) return;

    var formData = new FormData();
    for (var i = 0; i < files.length; i++) {
        var f = files[i];
        var name = f.name.toLowerCase();
        if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".pdf")) {
            formData.append("files", f, f.webkitRelativePath || f.name);
        }
    }
    folderInput.value = "";
    sendFiles(formData);
});

// ─── SHARED: Walk folder tree (for drag-drop) ──────────────

function walkTree(item, formData, path) {
    return new Promise(function(resolve) {
        if (item.isFile) {
            item.file(function(file) {
                var name = file.name.toLowerCase();
                if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".pdf")) {
                    formData.append("files", file, path + file.name);
                }
                resolve();
            });
        } else if (item.isDirectory) {
            var reader = item.createReader();
            readAllEntries(reader, function(entries) {
                var promises = [];
                for (var i = 0; i < entries.length; i++) {
                    promises.push(walkTree(entries[i], formData, path + item.name + "/"));
                }
                Promise.all(promises).then(resolve);
            });
        } else {
            resolve();
        }
    });
}

function readAllEntries(reader, callback) {
    var all = [];
    function batch() {
        reader.readEntries(function(entries) {
            if (entries.length === 0) { callback(all); }
            else { all = all.concat(Array.from(entries)); batch(); }
        });
    }
    batch();
}

// ─── SHARED: Send to server ─────────────────────────────────

async function sendFiles(formData) {
    fileList.innerHTML = "";
    summaryEl.style.display = "none";
    loader.style.display = "block";

    formData.append("month", document.getElementById("monthSel").value);
    formData.append("year", document.getElementById("yearIn").value);

    try {
        var resp = await fetch("/upload", { method: "POST", body: formData });
        var data = await resp.json();
        loader.style.display = "none";
        renderAll(data);
    } catch (err) {
        loader.style.display = "none";
        alert("Error: " + err.message);
    }
}

// ─── CLEAR ──────────────────────────────────────────────────

function clearAll() {
    fileList.innerHTML = "";
    summaryEl.style.display = "none";
    loader.style.display = "none";
}

// ─── RENDER ─────────────────────────────────────────────────

function renderAll(data) {
    if (data.error) { alert(data.error); return; }

    var results = data.results;
    var total = results.length;
    var ok = 0, fail = 0, matched = 0, wrong = 0;

    for (var i = 0; i < results.length; i++) {
        var r = results[i];
        if (r.success) { ok++; } else { fail++; }
        if (r.month_check) {
            if (r.month_check.status === "mismatch") { wrong++; }
            else if (r.month_check.match) { matched++; }
        }
    }

    document.getElementById("sTotal").textContent = total;
    document.getElementById("sOk").textContent    = ok;
    document.getElementById("sFail").textContent  = fail;
    document.getElementById("sMatch").textContent = matched;
    document.getElementById("sWrong").textContent = wrong;
    summaryEl.style.display = "block";

    results.sort(function(a, b) {
        var order = {"mismatch": 0, "warning": 1, "ok_multi": 2, "ok": 3};
        var sa = a.month_check ? (order[a.month_check.status] || 3) : 3;
        var sb = b.month_check ? (order[b.month_check.status] || 3) : 3;
        return sa - sb;
    });

    for (var i = 0; i < results.length; i++) {
        renderCard(results[i], i);
    }
}

function renderCard(res, idx) {
    var card = document.createElement("div");
    card.className = "file-card";

    if (res.month_check) {
        if (res.month_check.status === "mismatch") card.classList.add("mismatch");
        else if (res.month_check.status === "ok") card.classList.add("matched");
        else card.classList.add("warning");
    }

    var ext = res.filename.split(".").pop().toLowerCase();
    var extClass = ext === "pdf" ? "b-pdf" : ext === "xls" ? "b-xls" : "b-xlsx";
    var statusClass = res.success ? "b-ok" : "b-fail";
    var statusText  = res.success ? "OK" : "FAIL";
    var rowText = (res.total_rows || 0).toLocaleString() + " rows";
    var pvId = "pv_" + idx;
    var arId = "ar_" + idx;

    var monthBadge = "";
    if (res.month_check) {
        var mc = res.month_check;
        var mbClass;
        if (mc.status === "ok") { mbClass = "mb-match"; }
        else if (mc.status === "mismatch") { mbClass = "mb-mismatch"; }
        else { mbClass = "mb-warn"; }
        var shortMsg = mc.message.length > 60 ? mc.message.substring(0, 57) + "..." : mc.message;
        monthBadge = '<span class="month-badge ' + mbClass + '">' + mc.icon + " " + esc(shortMsg) + '</span>';
    }

    var header = document.createElement("div");
    header.className = "file-header";
    header.setAttribute("onclick", "toggle('" + pvId + "','" + arId + "')");
    header.innerHTML =
        '<div class="file-info">' +
            '<span class="badge ' + extClass + '">' + ext.toUpperCase() + '</span>' +
            '<b>' + esc(res.filename) + '</b>' +
            '<span class="badge ' + statusClass + '">' + statusText + '</span>' +
            monthBadge +
        '</div>' +
        '<div style="display:flex;align-items:center;gap:10px">' +
            '<span class="row-count">' + rowText + '</span>' +
            '<span class="arrow" id="' + arId + '">▼</span>' +
        '</div>';

    card.appendChild(header);

    if (res.month_check) {
        var mc = res.month_check;
        var barClass = mc.match ? "match" : mc.status === "mismatch" ? "mismatch" : "warn";
        var detStr = mc.detected.length > 0 ? "Detected: " + mc.detected.join(", ") : "No months detected";
        var bar = document.createElement("div");
        bar.className = "month-bar " + barClass;
        bar.innerHTML = '<span>' + mc.icon + ' ' + esc(mc.message) + '</span>' +
            '<span class="detected-list">' + esc(detStr) + '</span>';
        card.appendChild(bar);
    }

    var preview = document.createElement("div");
    preview.className = "preview-area";
    preview.id = pvId;

    if (!res.success) {
        preview.innerHTML = '<div class="error-preview">' + esc(res.error_message || "Unknown error") + '</div>';
    } else if (res.data_type === "table") {
        preview.innerHTML = buildTable(res.data_rows, res.data_cols);
    } else if (res.data_type === "text") {
        preview.innerHTML = buildText(res.data_lines);
    }

    card.appendChild(preview);
    fileList.appendChild(card);
}

function buildTable(rows, numCols) {
    var html = '<table class="data-table"><thead><tr><th class="row-num">#</th>';
    for (var c = 0; c < numCols; c++) {
        var letter;
        if (c < 26) { letter = String.fromCharCode(65 + c); }
        else { letter = String.fromCharCode(64 + Math.floor(c/26)) + String.fromCharCode(65 + (c % 26)); }
        html += '<th>' + letter + '</th>';
    }
    html += '</tr></thead><tbody>';
    for (var r = 0; r < rows.length; r++) {
        html += '<tr><td class="row-num">' + (r+1) + '</td>';
        for (var c = 0; c < numCols; c++) {
            var val = c < rows[r].length ? rows[r][c] : "";
            if (val === "" || val === "nan" || val === "None") {
                html += '<td class="cell-empty">&mdash;</td>';
            } else {
                html += '<td>' + esc(val) + '</td>';
            }
        }
        html += '</tr>';
    }
    html += '</tbody></table>';
    return html;
}

function buildText(lines) {
    var html = '<div class="text-preview">';
    for (var i = 0; i < lines.length; i++) {
        html += '<span class="line-num">' + (i+1) + '</span>' + esc(lines[i]) + '\n';
    }
    html += '</div>';
    return html;
}

function toggle(pvId, arId) {
    var el = document.getElementById(pvId);
    var ar = document.getElementById(arId);
    if (el.style.display === "block") {
        el.style.display = "none";
        ar.classList.remove("open");
    } else {
        el.style.display = "block";
        ar.classList.add("open");
    }
}

function esc(text) {
    var div = document.createElement("div");
    div.textContent = String(text);
    return div.innerHTML;
}
</script>
</body>
</html>
"""


# ─── ROUTES ──────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML_UI)


@app.route("/upload", methods=["POST"])
def upload():
    if "files" not in request.files:
        return jsonify({"error": "No files found"}), 400

    files = request.files.getlist("files")

    try:
        target_month = int(request.form.get("month", 1))
        target_year  = int(request.form.get("year", datetime.now().year))
    except ValueError:
        target_month = 1
        target_year  = datetime.now().year

    results = []

    for f in files:
        sid = str(uuid.uuid4())[:8]
        safe = sid + "_" + secure_filename(f.filename)
        filepath = UPLOAD_FOLDER / safe
        f.save(filepath)

        ext = filepath.suffix.lower()
        entry = {
            "filename": f.filename,
            "success": False,
            "total_rows": 0,
            "data_type": "error",
            "error_message": "",
            "month_check": None,
        }

        if ext in [".xlsx", ".xls"]:
            data = read_excel(filepath)
        elif ext == ".pdf":
            data = read_pdf(filepath)
        else:
            data = {"type": "error", "message": "Unsupported type: " + ext}

        if data["type"] == "error":
            entry["error_message"] = data.get("message", "Unknown error")
        elif data["type"] == "table":
            entry["success"]    = True
            entry["data_type"]  = "table"
            entry["data_rows"]  = data["rows"]
            entry["data_cols"]  = data["cols"]
            entry["total_rows"] = len(data["rows"])
            detected = detect_months_in_excel(data["rows"])
            detected.update(detect_months_in_text(f.filename))
            entry["month_check"] = validate_month(detected, target_year, target_month)
        elif data["type"] == "text":
            entry["success"]    = True
            entry["data_type"]  = "text"
            entry["data_lines"] = data["lines"]
            entry["total_rows"] = len(data["lines"])
            detected = detect_months_in_lines(data["lines"])
            detected.update(detect_months_in_text(f.filename))
            entry["month_check"] = validate_month(detected, target_year, target_month)

        results.append(entry)

        if filepath.exists():
            os.remove(filepath)

    return jsonify({"results": results})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("Starting on port", port)
    app.run(host="0.0.0.0", port=port, debug=False)
