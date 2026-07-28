# app.py
import os
import re
import uuid
import calendar as cal
from datetime import datetime, date, timedelta
import pandas as pd
import pdfplumber
from pathlib import Path
from flask import Flask, request, jsonify, render_template_string
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = Path("temp_uploads")
UPLOAD_FOLDER.mkdir(exist_ok=True)

# ─── INDONESIAN CALENDAR ─────────────────────────────────────
import holidays
def get_id_holidays(year): return set(holidays.Indonesia(years=year).keys())
def classify_day(d, hols):
    if d.weekday() >= 5 or d in hols: return "Weekend"
    return "Weekday"
def build_timeline(year, month):
    hols = get_id_holidays(year)
    first = date(year, month, 1)
    n = cal.monthrange(year, month)[1]
    return pd.DataFrame([{
        "Date": first + timedelta(days=i),
        "DayName": (first + timedelta(days=i)).strftime("%A"),
        "DayType": classify_day(first + timedelta(days=i), hols)
    } for i in range(n)])

# ─── MONTH DETECTION (Phase 1.5) ─────────────────────────────
MONTH_NAMES = {
    "jan":1,"january":1,"januari":1,"feb":2,"february":2,"februari":2,
    "mar":3,"march":3,"maret":3,"apr":4,"april":4,"may":5,"mei":5,
    "jun":6,"june":6,"juni":6,"jul":7,"july":7,"juli":7,"aug":8,
    "august":8,"agustus":8,"agu":8,"ags":8,"sep":9,"september":9,
    "sept":9,"oct":10,"october":10,"okt":10,"oktober":10,"nov":11,
    "november":11,"nop":11,"nopember":11,"dec":12,"december":12,
    "des":12,"desember":12
}
def detect_months_in_text(text):
    text = str(text).lower().strip()
    found = set()
    if not text or text in ("nan","none",""): return found
    clean = text.replace(",","").replace(".","").strip()
    if clean.isdigit(): return found
    for m in re.finditer(r"([a-z]{3,})[\s\-_/\.]+(\d{2,4})", text):
        name, yr = m.group(1), int(m.group(2))
        if yr < 100: yr += 2000
        mn = MONTH_NAMES.get(name)
        if mn and 2020 <= yr <= 2035: found.add((yr, mn))
    for m in re.finditer(r"(\d{4})[\s\-_/\.]+(\d{1,2})(?!\d)", text):
        yr, mn = int(m.group(1)), int(m.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035: found.add((yr, mn))
    for m in re.finditer(r"(\d{4})[\s\-_/\.]+([a-z]{3,})", text):
        yr = int(m.group(1))
        mn = MONTH_NAMES.get(m.group(2))
        if mn and 2020 <= yr <= 2035: found.add((yr, mn))
    for m in re.finditer(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", text):
        a,b,c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 2020 <= c <= 2035:
            if 1 <= b <= 12: found.add((c,b))
            if 1 <= a <= 12 and a != b: found.add((c,a))
    for m in re.finditer(r"(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})", text):
        yr, mn = int(m.group(1)), int(m.group(2))
        if 1 <= mn <= 12 and 2020 <= yr <= 2035: found.add((yr, mn))
    return found
def detect_months_in_excel(rows):
    found = set()
    for row in rows:
        for cell in row:
            s = str(cell).strip()
            if s and s not in ("","nan","None"): found.update(detect_months_in_text(s))
    return found
def validate_month(detected, target_year, target_month):
    target = (target_year, target_month)
    target_label = cal.month_name[target_month] + " " + str(target_year)
    if not detected:
        return {"status":"warning","icon":"⚠️","message":"No month/year detected. Cannot verify.","match":False,"detected":[],"target":target_label}
    detected_labels = sorted([cal.month_name[m]+" "+str(y) for y,m in detected])
    if target in detected:
        if len(detected)==1:
            return {"status":"ok","icon":"✅","message":"Matches: "+target_label,"match":True,"detected":detected_labels,"target":target_label}
        else:
            others = [l for l in detected_labels if l != target_label]
            return {"status":"ok_multi","icon":"⚠️","message":"Contains "+target_label+" but also has: "+", ".join(others),"match":True,"detected":detected_labels,"target":target_label}
    else:
        return {"status":"mismatch","icon":"⚠️","message":"WRONG MONTH — Expected "+target_label+" but found: "+", ".join(detected_labels),"match":False,"detected":detected_labels,"target":target_label}

# ─── FILE READERS ────────────────────────────────────────────
def read_excel(path):
    try:
        engine = "xlrd" if path.suffix == ".xls" else "openpyxl"
        df = pd.read_excel(path, header=None, engine=engine).fillna("")
        rows = [[str(c) if str(c)!="" else "" for c in row] for row in df.values.tolist()]
        return {"type":"table","rows":rows,"cols":len(rows[0]) if rows else 0}
    except Exception as e: return {"type":"error","message":"Excel Error: "+str(e)}

def smart_ocr_to_table(lines):
    table_rows = []
    date_pattern = re.compile(r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December|Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+(\d{4})", re.IGNORECASE)
    for line in lines:
        clean = str(line).strip()
        if not clean: continue
        merged = clean
        for day, month, year in date_pattern.findall(merged):
            original = day+" "+month+" "+year
            merged = merged.replace(original, day+"_"+month+"_"+year, 1)
        parts = re.split(r'\s+', merged)
        restored = [p.replace("_"," ") if "_" in p and date_pattern.match(p.replace("_"," ")) else p for p in parts]
        table_rows.append(restored)
    if not table_rows: return None, 0
    max_cols = max(len(row) for row in table_rows)
    for row in table_rows:
        while len(row) < max_cols: row.append("")
    return table_rows, max_cols

def read_pdf_ocr(path):
    try:
        from pdf2image import convert_from_path
        import pytesseract
        images = convert_from_path(path, dpi=300)
        text_lines = []
        for img in images:
            text = pytesseract.image_to_string(img, lang="eng")
            if text: text_lines.extend([l.strip() for l in text.split("\n") if l.strip()])
        if text_lines:
            rows, cols = smart_ocr_to_table(text_lines)
            if rows: return {"type":"table","rows":rows,"cols":cols,"is_ocr":True}
            return {"type":"text","lines":text_lines}
        return {"type":"error","message":"PDF OCR: No text detected."}
    except Exception as e: return {"type":"error","message":"PDF OCR Error: "+str(e)}

def read_pdf(path):
    try:
        tables_found, text_lines = [], []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    for table in page_tables:
                        for row in table:
                            cleaned = [str(c).strip() if c else "" for c in row]
                            if any(cleaned): tables_found.append(cleaned)
                text = page.extract_text()
                if text: text_lines.extend([l.strip() for l in text.split("\n") if l.strip()])
        if tables_found:
            max_cols = max(len(r) for r in tables_found)
            for r in tables_found:
                while len(r) < max_cols: r.append("")
            return {"type":"table","rows":tables_found,"cols":max_cols}
        if text_lines and len(text_lines) > 5:
            rows, cols = smart_ocr_to_table(text_lines)
            if rows: return {"type":"table","rows":rows,"cols":cols}
            return {"type":"text","lines":text_lines}
        return read_pdf_ocr(path)
    except Exception as e: return read_pdf_ocr(path)

# ─── PHASE 2: SMART DATA PARSER ──────────────────────────────
def sniff_columns(rows):
    """Find Date and Sales columns by analyzing content, not headers."""
    if not rows: return None, None
    num_cols = len(rows[0])
    date_scores = [0]*num_cols
    sales_scores = [0]*num_cols
    for row in rows[:15]:
        for i, cell in enumerate(row):
            if i >= num_cols: break
            val = str(cell).strip()
            if not val or val.lower() in ('nan','none',''): continue
            if re.search(r'\d{1,2}\s+[A-Za-z]+\s+\d{4}', val, re.I) or re.search(r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}', val):
                date_scores[i] += 2
            clean_num = re.sub(r'[Rp\s$€£]', '', val).replace(',', '')
            if clean_num.count('.') > 1: clean_num = clean_num.replace('.', '')
            elif clean_num.count('.') == 1:
                parts = clean_num.split('.')
                if len(parts[1]) == 3: clean_num = clean_num.replace('.', '')
            try:
                num = float(clean_num)
                if num > 10000: sales_scores[i] += 2
                elif num > 0: sales_scores[i] += 1
            except: pass
    date_col = date_scores.index(max(date_scores)) if max(date_scores) > 0 else None
    sales_col = sales_scores.index(max(sales_scores)) if max(sales_scores) > 0 else None
    if date_col == sales_col:
        if date_scores[date_col] > sales_scores[sales_col]: sales_col = None
        else: date_col = None
    return date_col, sales_col

def clean_indonesian_number(val):
    """Convert Indonesian/messy number strings to float."""
    s = str(val).strip()
    if not s or s.lower() in ('nan','none','-',''): return 0.0
    s = re.sub(r'[Rp\s$€£]', '', s).replace(',', '')
    if s.count('.') > 1: s = s.replace('.', '')
    elif s.count('.') == 1:
        parts = s.split('.')
        if len(parts[1]) == 3: s = s.replace('.', '')
    try: return float(s)
    except: return 0.0

def parse_date_flexible(val):
    """Try multiple date formats and return a date object or None."""
    s = str(val).strip().replace(',', '')
    for fmt in ['%d %B %Y','%d %b %Y','%d-%m-%Y','%d/%m/%Y','%Y-%m-%d','%d %B %Y']:
        try: return datetime.strptime(s, fmt).date()
        except: continue
    m = re.search(r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', s)
    if m:
        try: return datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", '%d %B %Y').date()
        except: pass
    return None

def parse_tenant_file(raw_rows, filename, target_year, target_month):
    """Extract unified daily sales data from messy rows."""
    if not raw_rows: return None, "No data rows found."
    
    # 1. Skip logos/titles/metadata: find first row with >=3 meaningful cells
    data_start = 0
    for i, row in enumerate(raw_rows):
        non_empty = [c for c in row if str(c).strip() and str(c).strip().lower() not in ('nan','none','')]
        if len(non_empty) >= 3:
            data_start = i
            break
            
    sample = raw_rows[data_start:data_start+20]
    date_col, sales_col = sniff_columns(sample)
    
    if date_col is None or sales_col is None:
        return None, f"Could not detect Date/Sales columns. (date_col={date_col}, sales_col={sales_col})"
        
    records = []
    for row in raw_rows[data_start:]:
        if len(row) <= max(date_col, sales_col): continue
        d = parse_date_flexible(row[date_col])
        if not d or d.year != target_year or d.month != target_month: continue
        sales = clean_indonesian_number(row[sales_col])
        records.append({'date': d, 'sales': sales})
        
    if not records: return None, "No valid records found for target month."
    
    df = pd.DataFrame(records)
    df = df.groupby('date')['sales'].max().reset_index() # Handle duplicates
    
    timeline = build_timeline(target_year, target_month)
    merged = timeline.merge(df, left_on='Date', right_on='date', how='left')
    merged['Sales'] = merged['sales'].fillna(0)
    merged.drop(columns=['date','sales'], inplace=True)
    
    tenant_name = Path(filename).stem.replace('_',' ').title()
    return {tenant_name: merged}, None

# ─── HTML UI ─────────────────────────────────────────────────
HTML_UI = r"""
<!DOCTYPE html>
<html>
<head>
    <title>Tenant Sales Aggregator</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <style>
        body{font-family:sans-serif;background:#0f172a;color:#e2e8f0;padding:40px;margin:0}
        .container{max-width:1200px;margin:0 auto}
        h1{margin-bottom:5px}.subtitle{color:#94a3b8;margin-bottom:25px}
        .config-card{background:#1e293b;border:1px solid #334155;border-radius:12px;padding:20px;margin-bottom:20px}
        .config-title{font-size:1rem;font-weight:600;color:#93c5fd;margin-bottom:15px}
        .config-row{display:flex;gap:15px;align-items:flex-end;flex-wrap:wrap}
        .config-group label{display:block;font-size:.8rem;color:#94a3b8;margin-bottom:5px}
        .config-group select,.config-group input[type=number]{background:#263248;border:1px solid #475569;border-radius:8px;color:#e2e8f0;padding:8px 12px;font-size:.95rem;outline:none}
        .config-group select option{background:#1e293b}
        .month-preview{font-size:.85rem;color:#94a3b8;padding:8px 0}.month-preview b{color:#3b82f6}
        .drop-zone{border:3px dashed #3b82f6;border-radius:20px;padding:50px;text-align:center;cursor:pointer;background:rgba(59,130,246,.05);transition:.3s}
        .drop-zone.hover{background:rgba(59,130,246,.15);border-color:#60a5fa;transform:scale(1.01)}
        .drop-zone .icon{font-size:3rem;margin-bottom:10px}.drop-zone p{color:#94a3b8;margin:5px 0}.drop-zone b{color:#93c5fd}
        .browse-row{display:flex;gap:10px;justify-content:center;margin-top:15px;flex-wrap:wrap}
        .browse-btn{display:inline-flex;align-items:center;gap:6px;padding:10px 20px;border-radius:10px;font-size:.9rem;font-weight:600;border:none;cursor:pointer;transition:all .2s}
        .browse-btn:hover{transform:translateY(-1px)}.btn-files{background:#3b82f6;color:#fff}.btn-files:hover{background:#2563eb}
        .btn-folder{background:#10b981;color:#fff}.btn-folder:hover{background:#059669}
        .btn-clear{background:transparent;color:#94a3b8;border:1px solid #475569}.btn-clear:hover{background:#263248}
        .hidden-input{display:none}
        .loader{display:none;text-align:center;color:#3b82f6;font-weight:bold;margin-top:15px;padding:15px}
        .spinner{display:inline-block;width:14px;height:14px;border:2px solid #3b82f6;border-top:2px solid transparent;border-radius:50%;animation:spin .6s linear infinite;margin-right:8px;vertical-align:middle}
        @keyframes spin{to{transform:rotate(360deg)}}
        .summary{margin-top:20px;display:none}.summary-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}
        .summary-card{background:#1e293b;border:1px solid #334155;padding:15px;border-radius:10px;text-align:center}
        .summary-card .num{font-size:1.5rem;font-weight:bold}.summary-card .label{font-size:.75rem;color:#94a3b8;margin-top:4px}
        .c-blue{color:#3b82f6}.c-green{color:#10b981}.c-red{color:#ef4444}.c-yellow{color:#f59e0b}.c-purple{color:#a78bfa}
        .file-list{margin-top:20px}.file-card{background:#1e293b;border:1px solid #334155;border-radius:12px;margin-bottom:12px;overflow:hidden}
        .file-card.mismatch{border-color:#ef4444}.file-card.matched{border-color:#10b981}.file-card.warning{border-color:#f59e0b}
        .file-header{display:flex;justify-content:space-between;align-items:center;padding:12px 15px;cursor:pointer;transition:background .2s}
        .file-header:hover{background:#263248}.file-info{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
        .badge{font-size:.7em;padding:3px 8px;border-radius:4px;font-weight:bold;letter-spacing:.5px}
        .b-xlsx{background:#2563eb;color:#fff}.b-xls{background:#7c3aed;color:#fff}.b-pdf{background:#f59e0b;color:#000}
        .b-ok{background:rgba(16,185,129,.2);color:#10b981;border:1px solid #10b981}.b-fail{background:rgba(239,68,68,.2);color:#ef4444;border:1px solid #ef4444}
        .month-badge{font-size:.72em;padding:3px 8px;border-radius:4px;font-weight:bold}
        .mb-match{background:rgba(16,185,129,.15);color:#10b981;border:1px solid rgba(16,185,129,.3)}
        .mb-mismatch{background:rgba(239,68,68,.15);color:#ef4444;border:1px solid rgba(239,68,68,.3)}
        .mb-warn{background:rgba(245,158,11,.15);color:#f59e0b;border:1px solid rgba(245,158,11,.3)}
        .row-count{font-size:.82em;color:#94a3b8}.arrow{color:#64748b;transition:transform .2s}.arrow.open{transform:rotate(180deg)}
        .month-bar{padding:8px 15px;font-size:.82em;display:flex;align-items:center;gap:8px}
        .month-bar.match{background:rgba(16,185,129,.08);color:#6ee7b7;border-top:1px solid rgba(16,185,129,.15)}
        .month-bar.mismatch{background:rgba(239,68,68,.08);color:#fca5a5;border-top:1px solid rgba(239,68,68,.15)}
        .month-bar.warn{background:rgba(245,158,11,.08);color:#fde68a;border-top:1px solid rgba(245,158,11,.15)}
        .month-bar .detected-list{font-size:.9em;color:#94a3b8;margin-left:auto}
        .preview-area{display:none;border-top:1px solid #334155;max-height:500px;overflow:auto;background:#0c1222}
        .data-table{width:100%;border-collapse:collapse;font-size:.8em;font-family:"Consolas","Courier New",monospace}
        .data-table th{background:#1a3a5c;color:#93c5fd;padding:6px 10px;text-align:left;font-weight:600;position:sticky;top:0;z-index:1;border-bottom:2px solid #2563eb;white-space:nowrap}
        .data-table td{padding:5px 10px;border-bottom:1px solid rgba(255,255,255,.05);color:#cbd5e1;white-space:nowrap}
        .data-table tr:hover td{background:rgba(59,130,246,.08)}.data-table .row-num{color:#475569;text-align:right;padding-right:12px;font-size:.85em;user-select:none;border-right:1px solid #334155;background:#111827}
        .data-table .cell-empty{color:#334155;font-style:italic}
        .text-preview{padding:15px;font-family:monospace;font-size:.8em;color:#10b981;white-space:pre-wrap;line-height:1.6}
        .text-preview .line-num{display:inline-block;width:40px;color:#475569;text-align:right;margin-right:12px;user-select:none}
        .error-preview{padding:15px;color:#ef4444;font-size:.85em}
        .process-btn{display:block;width:100%;padding:14px;margin-top:20px;background:linear-gradient(90deg,#3b82f6,#8b5cf6);color:#fff;border:none;border-radius:10px;font-size:1rem;font-weight:700;cursor:pointer;transition:.2s}
        .process-btn:hover{transform:translateY(-2px);box-shadow:0 4px 15px rgba(59,130,246,.4)}
        .process-btn:disabled{opacity:.4;cursor:not-allowed;transform:none;box-shadow:none}
        #unifiedSection{display:none;margin-top:30px}
        .unified-card{background:#1e293b;border:1px solid #334155;border-radius:12px;padding:20px;margin-bottom:20px}
        .chart-box{position:relative;height:350px;width:100%;margin-top:15px}
    </style>
</head>
<body>
<div class="container">
    <h1>📊 Tenant Sales Aggregator</h1>
    <p class="subtitle">Phase 2: Smart Parser & Unified Report Engine</p>

    <div class="config-card">
        <div class="config-title">📅 Target Report Month</div>
        <div class="config-row">
            <div class="config-group"><label>Month</label>
                <select id="monthSel"><option value="1">January</option><option value="2">February</option><option value="3">March</option><option value="4">April</option><option value="5">May</option><option value="6">June</option><option value="7">July</option><option value="8">August</option><option value="9">September</option><option value="10">October</option><option value="11">November</option><option value="12">December</option></select>
            </div>
            <div class="config-group"><label>Year</label><input id="yearIn" type="number" value="2026" min="2000" max="2100" style="width:90px"/></div>
            <div class="month-preview">Target: <b id="targetLabel">January 2026</b></div>
        </div>
    </div>

    <div class="drop-zone" id="dropZone">
        <div class="icon">📥</div><p><b>Drop files or folder here</b></p><p style="font-size:.85em">Or use buttons below</p>
    </div>
    <div class="browse-row">
        <button class="browse-btn btn-files" onclick="document.getElementById('fileInput').click()">📄 Browse Files</button>
        <button class="browse-btn btn-folder" onclick="document.getElementById('folderInput').click()">📁 Browse Folder</button>
        <button class="browse-btn btn-clear" onclick="clearAll()">↺ Clear All</button>
    </div>
    <input type="file" id="fileInput" class="hidden-input" accept=".xlsx,.xls,.pdf" multiple/>
    <input type="file" id="folderInput" class="hidden-input" webkitdirectory/>

    <div id="loader" class="loader"><span class="spinner"></span> Reading & validating files...</div>
    <div class="summary" id="summary">
        <div class="summary-grid">
            <div class="summary-card"><div class="num c-blue" id="sTotal">0</div><div class="label">Total Files</div></div>
            <div class="summary-card"><div class="num c-green" id="sOk">0</div><div class="label">Read OK</div></div>
            <div class="summary-card"><div class="num c-red" id="sFail">0</div><div class="label">Read Failed</div></div>
            <div class="summary-card"><div class="num c-purple" id="sMatch">0</div><div class="label">Month Match ✅</div></div>
            <div class="summary-card"><div class="num c-yellow" id="sWrong">0</div><div class="label">Wrong Month ❌</div></div>
        </div>
    </div>
    <div class="file-list" id="fileList"></div>

    <button class="process-btn" id="processBtn" disabled onclick="processFiles()">⚙️ Process & Unify Data</button>

    <div id="unifiedSection">
        <div class="unified-card">
            <div class="config-title">📋 Unified Monthly Summary</div>
            <div class="summary-grid" id="unifiedStats"></div>
            <div style="overflow-x:auto;border-radius:10px;border:1px solid #334155;margin-top:15px">
                <table class="data-table" id="unifiedTable"><thead></thead><tbody></tbody></table>
            </div>
        </div>
        <div class="unified-card">
            <div class="config-title">📈 Daily Sales Trends</div>
            <div class="chart-box"><canvas id="unifiedChart"></canvas></div>
        </div>
    </div>
</div>

<script>
var dropZone=document.getElementById("dropZone"),fileList=document.getElementById("fileList"),loader=document.getElementById("loader"),summaryEl=document.getElementById("summary"),fileInput=document.getElementById("fileInput"),folderInput=document.getElementById("folderInput"),processBtn=document.getElementById("processBtn");
function updateLabel(){var m=["","January","February","March","April","May","June","July","August","September","October","November","December"];document.getElementById("targetLabel").textContent=m[parseInt(document.getElementById("monthSel").value)]+" "+document.getElementById("yearIn").value}
document.getElementById("monthSel").addEventListener("change",updateLabel);document.getElementById("yearIn").addEventListener("input",updateLabel);
var now=new Date();document.getElementById("monthSel").value=now.getMonth()+1;document.getElementById("yearIn").value=now.getFullYear();updateLabel();
dropZone.addEventListener("dragover",function(e){e.preventDefault();dropZone.classList.add("hover")});dropZone.addEventListener("dragleave",function(){dropZone.classList.remove("hover")});
dropZone.addEventListener("drop",async function(e){e.preventDefault();dropZone.classList.remove("hover");var items=e.dataTransfer.items,formData=new FormData();for(var i=0;i<items.length;i++){var entry=items[i].webkitGetAsEntry();if(entry)await walkTree(entry,formData,"")}sendFiles(formData)});
fileInput.addEventListener("change",function(){var files=fileInput.files,formData=new FormData();for(var i=0;i<files.length;i++){var f=files[i],n=f.name.toLowerCase();if(n.endsWith(".xlsx")||n.endsWith(".xls")||n.endsWith(".pdf"))formData.append("files",f,f.name)}fileInput.value="";sendFiles(formData)});
folderInput.addEventListener("change",function(){var files=folderInput.files,formData=new FormData();for(var i=0;i<files.length;i++){var f=files[i],n=f.name.toLowerCase();if(n.endsWith(".xlsx")||n.endsWith(".xls")||n.endsWith(".pdf"))formData.append("files",f,f.webkitRelativePath||f.name)}folderInput.value="";sendFiles(formData)});
function walkTree(item,formData,path){return new Promise(function(resolve){if(item.isFile)item.file(function(file){var n=file.name.toLowerCase();if(n.endsWith(".xlsx")||n.endsWith(".xls")||n.endsWith(".pdf"))formData.append("files",file,path+file.name);resolve()});else if(item.isDirectory){var reader=item.createReader();readAllEntries(reader,function(entries){var promises=[];for(var i=0;i<entries.length;i++)promises.push(walkTree(entries[i],formData,path+item.name+"/"));Promise.all(promises).then(resolve)})}else resolve()})}
function readAllEntries(reader,callback){var all=[];function batch(){reader.readEntries(function(entries){if(entries.length===0)callback(all);else{all=all.concat(Array.from(entries));batch()}})}batch()}
async function sendFiles(formData){fileList.innerHTML="";summaryEl.style.display="none";loader.style.display="block";formData.append("month",document.getElementById("monthSel").value);formData.append("year",document.getElementById("yearIn").value);try{var resp=await fetch("/upload",{method:"POST",body:formData}),data=await resp.json();loader.style.display="none";renderAll(data)}catch(err){loader.style.display="none";alert("Error: "+err.message)}}
function clearAll(){fileList.innerHTML="";summaryEl.style.display="none";loader.style.display="none";document.getElementById("unifiedSection").style.display="none";processBtn.disabled=true}
function renderAll(data){if(data.error){alert(data.error);return}var results=data.results,total=results.length,ok=0,fail=0,matched=0,wrong=0;for(var i=0;i<results.length;i++){var r=results[i];if(r.success)ok++;else fail++;if(r.month_check){if(r.month_check.status==="mismatch")wrong++;else if(r.month_check.status==="ok")matched++}}document.getElementById("sTotal").textContent=total;document.getElementById("sOk").textContent=ok;document.getElementById("sFail").textContent=fail;document.getElementById("sMatch").textContent=matched;document.getElementById("sWrong").textContent=wrong;summaryEl.style.display="block";results.sort(function(a,b){var order={"mismatch":0,"warning":1,"ok_multi":2,"ok":3};var sa=a.month_check?(order[a.month_check.status]||3):3,sb=b.month_check?(order[b.month_check.status]||3):3;return sa-sb});for(var i=0;i<results.length;i++)renderCard(results[i],i);processBtn.disabled=results.length===0}
function renderCard(res,idx){var card=document.createElement("div");card.className="file-card";if(res.month_check){if(res.month_check.status==="mismatch")card.classList.add("mismatch");else if(res.month_check.status==="ok")card.classList.add("matched");else card.classList.add("warning")}var ext=res.filename.split(".").pop().toLowerCase(),extClass=ext==="pdf"?"b-pdf":ext==="xls"?"b-xls":"b-xlsx",statusClass=res.success?"b-ok":"b-fail",statusText=res.success?"OK":"FAIL",rowText=(res.total_rows||0).toLocaleString()+" rows",pvId="pv_"+idx,arId="ar_"+idx,monthBadge="";if(res.month_check){var mc=res.month_check,mbClass=mc.status==="ok"?"mb-match":mc.status==="mismatch"?"mb-mismatch":"mb-warn",shortMsg=mc.message.length>60?mc.message.substring(0,57)+"...":mc.message;monthBadge='<span class="month-badge '+mbClass+'">'+mc.icon+" "+esc(shortMsg)+'</span>'}var header=document.createElement("div");header.className="file-header";header.setAttribute("onclick","toggle('"+pvId+"','"+arId+"')");header.innerHTML='<div class="file-info"><span class="badge '+extClass+'">'+ext.toUpperCase()+'</span><b>'+esc(res.filename)+'</b><span class="badge '+statusClass+'">'+statusText+'</span>'+monthBadge+'</div><div style="display:flex;align-items:center;gap:10px"><span class="row-count">'+rowText+'</span><span class="arrow" id="'+arId+'">▼</span></div>';card.appendChild(header);if(res.month_check){var mc=res.month_check,barClass=mc.status==="ok"?"match":mc.status==="mismatch"?"mismatch":"warn",detStr=mc.detected.length>0?"Detected: "+mc.detected.join(", "):"No months detected",bar=document.createElement("div");bar.className="month-bar "+barClass;bar.innerHTML='<span>'+mc.icon+' '+esc(mc.message)+'</span><span class="detected-list">'+esc(detStr)+'</span>';card.appendChild(bar)}var preview=document.createElement("div");preview.className="preview-area";preview.id=pvId;if(!res.success)preview.innerHTML='<div class="error-preview">'+esc(res.error_message||"Unknown error")+'</div>';else if(res.data_type==="table")preview.innerHTML=buildTable(res.data_rows,res.data_cols);else if(res.data_type==="text")preview.innerHTML=buildText(res.data_lines);card.appendChild(preview);fileList.appendChild(card)}
function buildTable(rows,numCols){var html='<table class="data-table"><thead><tr><th class="row-num">#</th>';for(var c=0;c<numCols;c++){var letter;c<26?letter=String.fromCharCode(65+c):letter=String.fromCharCode(64+Math.floor(c/26))+String.fromCharCode(65+(c%26));html+='<th>'+letter+'</th>'}html+='</tr></thead><tbody>';for(var r=0;r<rows.length;r++){html+='<tr><td class="row-num">'+(r+1)+'</td>';for(var c=0;c<numCols;c++){var val=c<rows[r].length?rows[r][c]:"";if(val===""||val==="nan"||val==="None")html+='<td class="cell-empty">&mdash;</td>';else html+='<td>'+esc(val)+'</td>'}html+='</tr>'}html+='</tbody></table>';return html}
function buildText(lines){var html='<div class="text-preview">';for(var i=0;i<lines.length;i++)html+='<span class="line-num">'+(i+1)+'</span>'+esc(lines[i])+'\n';html+='</div>';return html}
function toggle(pvId,arId){var el=document.getElementById(pvId),ar=document.getElementById(arId);if(el.style.display==="block"){el.style.display="none";ar.classList.remove("open")}else{el.style.display="block";ar.classList.add("open")}}
function esc(text){var div=document.createElement("div");div.textContent=String(text);return div.innerHTML}

// ─── PHASE 2: PROCESS & UNIFY ───────────────────────────────
async function processFiles() {
    processBtn.disabled = true;
    processBtn.textContent = "⏳ Processing & Unifying Data...";
    document.getElementById("unifiedSection").style.display = "none";

    var formData = new FormData();
    formData.append("month", document.getElementById("monthSel").value);
    formData.append("year", document.getElementById("yearIn").value);

    // Re-send files for backend parsing
    var items = dropZone.querySelectorAll ? [] : []; // We'll rely on session or re-upload
    // Simpler: trigger backend process route
    try {
        var resp = await fetch("/process", {method:"POST", body: formData});
        var data = await resp.json();
        if(data.error) { alert(data.error); return; }
        renderUnified(data);
    } catch(e) { alert("Processing error: "+e.message); }
    finally {
        processBtn.disabled = false;
        processBtn.textContent = "⚙️ Process & Unify Data";
    }
}

function renderUnified(data) {
    document.getElementById("unifiedSection").style.display = "block";
    var stats = document.getElementById("unifiedStats");
    stats.innerHTML = `
        <div class="summary-card"><div class="num c-blue">${data.tenant_count}</div><div class="label">Tenants Parsed</div></div>
        <div class="summary-card"><div class="num c-green">${data.total_rows}</div><div class="label">Total Daily Rows</div></div>
        <div class="summary-card"><div class="num c-yellow">${data.missing_dates}</div><div class="label">Dates Auto-Filled (0)</div></div>
    `;

    var thead = document.querySelector("#unifiedTable thead");
    var tbody = document.querySelector("#unifiedTable tbody");
    thead.innerHTML = "<tr><th>Tenant</th><th>Total Sales</th><th>Weekday Avg</th><th>Weekend Avg</th><th>Days w/ Sales</th></tr>";
    tbody.innerHTML = "";
    var grandTotal = 0;
    data.summary.forEach(function(r) {
        grandTotal += r.total;
        tbody.innerHTML += `<tr><td><strong>${r.tenant}</strong></td><td>${fmtIDR(r.total)}</td><td>${fmtIDR(r.wd_avg)}</td><td>${fmtIDR(r.we_avg)}</td><td>${r.active_days}</td></tr>`;
    });
    tbody.innerHTML += `<tr style="background:rgba(59,130,246,.08);font-weight:700"><td>🏆 Grand Total</td><td>${fmtIDR(grandTotal)}</td><td colspan="3" style="text-align:center;color:#94a3b8">— all tenants combined —</td></tr>`;

    // Chart
    var ctx = document.getElementById("unifiedChart").getContext("2d");
    if(window.unifiedChartInstance) window.unifiedChartInstance.destroy();
    var colors = ["#3b82f6","#10b981","#f59e0b","#ef4444","#8b5cf6","#ec4899","#06b6d4","#84cc16"];
    var datasets = Object.entries(data.charts).map(function(entry, i) {
        return {
            label: entry[0],
            data: entry[1].values,
            borderColor: colors[i % colors.length],
            backgroundColor: colors[i % colors.length] + "22",
            borderWidth: 2, pointRadius: 3, tension: 0.3, fill: false
        };
    });
    window.unifiedChartInstance = new Chart(ctx, {
        type: "line",
        data: { labels: data.charts[Object.keys(data.charts)[0]].labels, datasets: datasets },
        options: {
            responsive: true, maintainAspectRatio: false,
            interaction: { mode: "index", intersect: false },
            plugins: { legend: { labels: { color: "#94a3b8" } }, tooltip: { backgroundColor: "#1e293b", titleColor: "#93c5fd", bodyColor: "#e2e8f0", callbacks: { label: function(ctx) { return " "+ctx.dataset.label+": "+fmtIDR(ctx.parsed.y); } } } },
            scales: { x: { ticks: { color: "#64748b", maxRotation: 45 }, grid: { color: "rgba(255,255,255,.04)" } }, y: { ticks: { color: "#64748b", callback: function(v){return fmtShort(v)} }, grid: { color: "rgba(255,255,255,.06)" } } }
        }
    });
    document.getElementById("unifiedSection").scrollIntoView({behavior:"smooth"});
}

function fmtIDR(v){return Number(v).toLocaleString("id-ID")}
function fmtShort(v){if(v>=1e12)return(v/1e12).toFixed(1)+"T";if(v>=1e9)return(v/1e9).toFixed(2)+"B";if(v>=1e6)return(v/1e6).toFixed(1)+"M";if(v>=1e3)return(v/1e3).toFixed(0)+"K";return v}
</script>
</body>
</html>
"""

# ─── FLASK ROUTES ────────────────────────────────────────────
@app.route("/")
def index(): return render_template_string(HTML_UI)

@app.route("/upload", methods=["POST"])
def upload():
    if "files" not in request.files: return jsonify({"error":"No files found"}), 400
    files = request.files.getlist("files")
    try: target_month=int(request.form.get("month",1)); target_year=int(request.form.get("year",datetime.now().year))
    except: target_month=1; target_year=datetime.now().year
    results = []
    for f in files:
        sid=str(uuid.uuid4())[:8]; safe=sid+"_"+secure_filename(f.filename); filepath=UPLOAD_FOLDER/safe; f.save(filepath)
        ext=filepath.suffix.lower(); entry={"filename":f.filename,"success":False,"total_rows":0,"data_type":"error","error_message":"","month_check":None}
        if ext in [".xlsx",".xls"]: data=read_excel(filepath)
        elif ext==".pdf": data=read_pdf(filepath)
        else: data={"type":"error","message":"Unsupported type: "+ext}
        if data["type"]=="error": entry["error_message"]=data.get("message","Unknown error")
        elif data["type"]=="table":
            entry["success"]=True; entry["data_type"]="table"; entry["data_rows"]=data["rows"]; entry["data_cols"]=data["cols"]; entry["total_rows"]=len(data["rows"])
            detected=detect_months_in_excel(data["rows"]); detected.update(detect_months_in_text(f.filename)); entry["month_check"]=validate_month(detected,target_year,target_month)
        elif data["type"]=="text":
            entry["success"]=True; entry["data_type"]="text"; entry["data_lines"]=data["lines"]; entry["total_rows"]=len(data["lines"])
            detected=detect_months_in_lines(data["lines"]); detected.update(detect_months_in_text(f.filename)); entry["month_check"]=validate_month(detected,target_year,target_month)
        results.append(entry)
        if filepath.exists(): os.remove(filepath)
    return jsonify({"results":results})

@app.route("/process", methods=["POST"])
def process():
    try: target_month=int(request.form.get("month",1)); target_year=int(request.form.get("year",datetime.now().year))
    except: target_month=1; target_year=datetime.now().year
    
    # Re-read files from last upload session (simplified: we'll re-use upload logic in production)
    # For now, we expect files to be re-sent or stored. We'll store them temporarily.
    # To keep it simple: we'll process files saved in UPLOAD_FOLDER from the last /upload call
    # Actually, let's just process files directly from a fresh upload in /process for reliability
    if "files" not in request.files:
        # Fallback: process already uploaded files if any
        files = list(UPLOAD_FOLDER.glob("*"))
        if not files: return jsonify({"error":"No files found. Please upload first."}), 400
    else:
        files = request.files.getlist("files")
        for f in files:
            safe = secure_filename(f.filename)
            f.save(UPLOAD_FOLDER / safe)
        files = list(UPLOAD_FOLDER.glob("*"))

    all_tenants = {}
    errors = []
    for fp in files:
        if fp.suffix.lower() not in [".xlsx",".xls",".pdf"]: continue
        if fp.suffix.lower() in [".xlsx",".xls"]: data = read_excel(fp)
        else: data = read_pdf(fp)
        
        if data["type"] != "table": 
            errors.append(f"{fp.name}: {data.get('message','No table data')}")
            continue
            
        result, err = parse_tenant_file(data["rows"], fp.name, target_year, target_month)
        if err: errors.append(f"{fp.name}: {err}")
        else: all_tenants.update(result)
        
        if fp.exists(): os.remove(fp)

    if not all_tenants:
        return jsonify({"error": "No valid tenant data parsed. Details: " + "; ".join(errors)}), 422

    # Build unified payload
    summary = []
    charts = {}
    total_rows = 0
    missing_dates = 0
    first_labels = None

    for tenant, df in all_tenants.items():
        total = float(df["Sales"].sum())
        wd = df[df["DayType"]=="Weekday"]["Sales"]
        we = df[df["DayType"]=="Weekend"]["Sales"]
        wd_avg = float(wd.mean()) if len(wd) else 0
        we_avg = float(we.mean()) if len(we) else 0
        active = int(df["Sales"].gt(0).sum())
        missing_dates += int(df["Sales"].eq(0).sum())
        total_rows += len(df)
        
        summary.append({"tenant":tenant,"total":total,"wd_avg":round(wd_avg,0),"we_avg":round(we_avg,0),"active_days":active})
        charts[tenant] = {"labels":[str(d) for d in df["Date"]],"values":[float(v) for v in df["Sales"]]}
        if not first_labels: first_labels = charts[tenant]["labels"]

    return jsonify({
        "tenant_count": len(all_tenants),
        "total_rows": total_rows,
        "missing_dates": missing_dates,
        "summary": summary,
        "charts": charts
    })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("Starting on port", port)
    app.run(host="0.0.0.0", port=port, debug=False)
