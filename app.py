# app.py
import os
import uuid
import pandas as pd
import pdfplumber
from pathlib import Path
from flask import Flask, request, jsonify, render_template_string
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = Path("temp_uploads")
UPLOAD_FOLDER.mkdir(exist_ok=True)


def read_excel(path):
    try:
        engine = "xlrd" if path.suffix == ".xls" else "openpyxl"
        df = pd.read_excel(path, header=None, engine=engine).fillna("")
        rows = df.values.tolist()
        # Convert every cell to string for safe JSON
        cleaned = []
        for row in rows:
            cleaned.append([str(cell) if str(cell) != "" else "" for cell in row])
        return {"type": "table", "rows": cleaned, "cols": len(cleaned[0]) if cleaned else 0}
    except Exception as e:
        return {"type": "error", "message": f"Excel Error: {str(e)}"}


def read_pdf(path):
    try:
        # First try to extract tables
        tables_found = []
        text_lines = []

        with pdfplumber.open(path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # Try table extraction first
                page_tables = page.extract_tables()
                if page_tables:
                    for table in page_tables:
                        for row in table:
                            cleaned_row = [str(cell).strip() if cell else "" for cell in row]
                            tables_found.append(cleaned_row)

                # Also get raw text as fallback
                text = page.extract_text()
                if text:
                    text_lines.extend([l.strip() for l in text.split("\n") if l.strip()])

        # If we found tables, return as table format
        if tables_found:
            max_cols = max(len(row) for row in tables_found)
            # Pad rows to same length
            for row in tables_found:
                while len(row) < max_cols:
                    row.append("")
            return {"type": "table", "rows": tables_found, "cols": max_cols}

        # Otherwise return as text lines
        if text_lines:
            return {"type": "text", "lines": text_lines}

        return {"type": "error", "message": "PDF: No text or tables found."}
    except Exception as e:
        return {"type": "error", "message": f"PDF Error: {str(e)}"}


HTML_UI = r"""
<!DOCTYPE html>
<html>
<head>
    <title>Phase 1: Folder Reader</title>
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

        /* Summary Bar */
        .summary {
            margin-top: 20px;
            display: none;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
        }
        .summary-card {
            background: #1e293b;
            border: 1px solid #334155;
            padding: 15px;
            border-radius: 10px;
            text-align: center;
        }
        .summary-card .num {
            font-size: 1.6rem;
            font-weight: bold;
        }
        .summary-card .label {
            font-size: 0.78rem;
            color: #94a3b8;
            margin-top: 4px;
        }
        .c-blue { color: #3b82f6; }
        .c-green { color: #10b981; }
        .c-red { color: #ef4444; }
        .c-yellow { color: #f59e0b; }

        /* File Cards */
        .file-list { margin-top: 20px; }
        .file-card {
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 12px;
            margin-bottom: 12px;
            overflow: hidden;
        }
        .file-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 12px 15px;
            cursor: pointer;
            transition: background 0.2s;
        }
        .file-header:hover { background: #263248; }
        .file-info { display: flex; align-items: center; gap: 10px; }

        /* Badges */
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

        .row-count { font-size: 0.82em; color: #94a3b8; }
        .arrow { color: #64748b; transition: transform 0.2s; }
        .arrow.open { transform: rotate(180deg); }

        /* Preview Area */
        .preview-area {
            display: none;
            border-top: 1px solid #334155;
            max-height: 500px;
            overflow: auto;
            background: #0c1222;
        }

        /* Table Preview */
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
        .data-table tr:hover td {
            background: rgba(59, 130, 246, 0.08);
        }
        .data-table .row-num {
            color: #475569;
            text-align: right;
            padding-right: 12px;
            font-size: 0.85em;
            user-select: none;
            border-right: 1px solid #334155;
            background: #111827;
        }
        .data-table .cell-empty {
            color: #334155;
            font-style: italic;
        }

        /* Text Preview (for PDFs without tables) */
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

        /* Error Preview */
        .error-preview {
            padding: 15px;
            color: #ef4444;
            font-size: 0.85em;
        }
    </style>
</head>
<body>
<div class="container">
    <h1>📂 Tenant Report Reader</h1>
    <p class="subtitle">Drag and drop a folder containing tenant Excel or PDF reports.</p>

    <div class="drop-zone" id="dropZone">
        <div class="icon">📥</div>
        <p><b>Drop folder here</b></p>
        <p style="font-size: 0.85em;">Accepts .xlsx, .xls, and .pdf files</p>
    </div>

    <div id="loader" class="loader">
        <span class="spinner"></span> Reading files, please wait...
    </div>

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
                <div class="label">Failed</div>
            </div>
            <div class="summary-card">
                <div class="num c-yellow" id="sRows">0</div>
                <div class="label">Total Rows</div>
            </div>
        </div>
    </div>

    <div class="file-list" id="fileList"></div>
</div>

<script>
var dropZone = document.getElementById("dropZone");
var fileList = document.getElementById("fileList");
var loader   = document.getElementById("loader");
var summary  = document.getElementById("summary");

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
    fileList.innerHTML = "";
    summary.style.display = "none";
    loader.style.display = "block";

    var items = e.dataTransfer.items;
    var formData = new FormData();

    for (var i = 0; i < items.length; i++) {
        var item = items[i].webkitGetAsEntry();
        if (item) {
            await walkTree(item, formData, "");
        }
    }

    try {
        var resp = await fetch("/upload", { method: "POST", body: formData });
        var data = await resp.json();
        loader.style.display = "none";
        renderAll(data);
    } catch (err) {
        loader.style.display = "none";
        alert("Error: " + err.message);
    }
});

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
    var allEntries = [];
    function readBatch() {
        reader.readEntries(function(entries) {
            if (entries.length === 0) {
                callback(allEntries);
            } else {
                allEntries = allEntries.concat(Array.from(entries));
                readBatch();
            }
        });
    }
    readBatch();
}

function renderAll(data) {
    if (data.error) { alert(data.error); return; }

    var results = data.results;
    var total = results.length;
    var ok = 0;
    var fail = 0;
    var rows = 0;

    for (var i = 0; i < results.length; i++) {
        if (results[i].success) { ok++; } else { fail++; }
        rows += results[i].total_rows || 0;
    }

    document.getElementById("sTotal").textContent = total;
    document.getElementById("sOk").textContent    = ok;
    document.getElementById("sFail").textContent   = fail;
    document.getElementById("sRows").textContent   = rows.toLocaleString();
    summary.style.display = "block";

    for (var i = 0; i < results.length; i++) {
        renderFileCard(results[i], i);
    }
}

function renderFileCard(res, idx) {
    var card = document.createElement("div");
    card.className = "file-card";

    var ext = res.filename.split(".").pop().toLowerCase();
    var extClass = ext === "pdf" ? "b-pdf" : ext === "xls" ? "b-xls" : "b-xlsx";
    var statusClass = res.success ? "b-ok" : "b-fail";
    var statusText  = res.success ? "OK" : "FAIL";
    var rowText = (res.total_rows || 0).toLocaleString() + " rows";
    var previewId = "pv_" + idx;
    var arrowId   = "ar_" + idx;

    var header = document.createElement("div");
    header.className = "file-header";
    header.setAttribute("onclick", "toggleCard('" + previewId + "','" + arrowId + "')");
    header.innerHTML =
        '<div class="file-info">' +
            '<span class="badge ' + extClass + '">' + ext.toUpperCase() + '</span>' +
            '<b>' + esc(res.filename) + '</b>' +
            '<span class="badge ' + statusClass + '">' + statusText + '</span>' +
        '</div>' +
        '<div style="display:flex;align-items:center;gap:10px">' +
            '<span class="row-count">' + rowText + '</span>' +
            '<span class="arrow" id="' + arrowId + '">▼</span>' +
        '</div>';

    var previewArea = document.createElement("div");
    previewArea.className = "preview-area";
    previewArea.id = previewId;

    if (!res.success) {
        previewArea.innerHTML = '<div class="error-preview">❌ ' + esc(res.error_message || "Unknown error") + '</div>';
    } else if (res.data_type === "table") {
        previewArea.innerHTML = buildTable(res.data_rows, res.data_cols);
    } else if (res.data_type === "text") {
        previewArea.innerHTML = buildText(res.data_lines);
    }

    card.appendChild(header);
    card.appendChild(previewArea);
    fileList.appendChild(card);
}

function buildTable(rows, numCols) {
    var html = '<table class="data-table"><thead><tr>';
    html += '<th class="row-num">#</th>';
    for (var c = 0; c < numCols; c++) {
        html += '<th>Col ' + String.fromCharCode(65 + (c % 26)) + (c >= 26 ? String.fromCharCode(65 + Math.floor(c/26) - 1) : '') + '</th>';
    }
    html += '</tr></thead><tbody>';

    for (var r = 0; r < rows.length; r++) {
        html += '<tr>';
        html += '<td class="row-num">' + (r + 1) + '</td>';
        for (var c = 0; c < numCols; c++) {
            var val = c < rows[r].length ? rows[r][c] : "";
            if (val === "" || val === "nan" || val === "None") {
                html += '<td class="cell-empty">—</td>';
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
        html += '<span class="line-num">' + (i + 1) + '</span>' + esc(lines[i]) + '\n';
    }
    html += '</div>';
    return html;
}

function toggleCard(previewId, arrowId) {
    var el = document.getElementById(previewId);
    var ar = document.getElementById(arrowId);
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


@app.route("/")
def index():
    return render_template_string(HTML_UI)


@app.route("/upload", methods=["POST"])
def upload():
    if "files" not in request.files:
        return jsonify({"error": "No files found"}), 400

    files = request.files.getlist("files")
    results = []

    for f in files:
        sid = str(uuid.uuid4())[:8]
        safe = sid + "_" + secure_filename(f.filename)
        filepath = UPLOAD_FOLDER / safe
        f.save(filepath)

        ext = filepath.suffix.lower()
        result_entry = {
            "filename": f.filename,
            "success": False,
            "total_rows": 0,
            "data_type": "error",
            "error_message": "",
        }

        if ext in [".xlsx", ".xls"]:
            data = read_excel(filepath)
        elif ext == ".pdf":
            data = read_pdf(filepath)
        else:
            data = {"type": "error", "message": "Unsupported file type: " + ext}

        if data["type"] == "error":
            result_entry["error_message"] = data.get("message", "Unknown error")
        elif data["type"] == "table":
            result_entry["success"]    = True
            result_entry["data_type"]  = "table"
            result_entry["data_rows"]  = data["rows"]
            result_entry["data_cols"]  = data["cols"]
            result_entry["total_rows"] = len(data["rows"])
        elif data["type"] == "text":
            result_entry["success"]    = True
            result_entry["data_type"]  = "text"
            result_entry["data_lines"] = data["lines"]
            result_entry["total_rows"] = len(data["lines"])

        results.append(result_entry)

        if filepath.exists():
            os.remove(filepath)

    return jsonify({"results": results})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting on port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)
