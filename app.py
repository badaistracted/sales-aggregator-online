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
        return df.values.tolist()
    except Exception as e:
        return f"Excel Error: {str(e)}"


def read_pdf(path):
    try:
        lines = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines.extend([l.strip() for l in text.split("\n") if l.strip()])
        return lines if lines else "PDF Error: No text found."
    except Exception as e:
        return f"PDF Error: {str(e)}"


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
        }
        .container { max-width: 1000px; margin: 0 auto; }

        .drop-zone {
            border: 3px dashed #3b82f6;
            border-radius: 20px;
            padding: 60px;
            text-align: center;
            cursor: pointer;
            background: rgba(59, 130, 246, 0.05);
            transition: 0.3s;
        }
        .drop-zone.hover {
            background: rgba(59, 130, 246, 0.2);
            border-color: #60a5fa;
        }

        .summary {
            margin-top: 20px;
            padding: 15px;
            background: #1e293b;
            border-radius: 10px;
            border: 1px solid #334155;
            display: none;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 10px;
            margin-top: 10px;
        }
        .summary-card {
            background: #263248;
            padding: 12px;
            border-radius: 8px;
            text-align: center;
        }
        .summary-card .num {
            font-size: 1.8rem;
            font-weight: bold;
            color: #3b82f6;
        }
        .summary-card .label {
            font-size: 0.8rem;
            color: #94a3b8;
            margin-top: 4px;
        }

        .file-list { margin-top: 20px; }

        .file-card {
            background: #1e293b;
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 12px;
            border: 1px solid #334155;
        }
        .file-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            cursor: pointer;
        }
        .file-info { display: flex; align-items: center; gap: 10px; }

        .badge {
            font-size: 0.75em;
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: bold;
        }
        .badge.ok { background: #10b981; color: #000; }
        .badge.fail { background: #ef4444; color: #fff; }
        .badge.xlsx { background: #2563eb; color: #fff; }
        .badge.xls { background: #7c3aed; color: #fff; }
        .badge.pdf { background: #f59e0b; color: #000; }

        .row-count {
            font-size: 0.85em;
            color: #94a3b8;
            font-weight: bold;
        }

        .preview {
            font-family: monospace;
            font-size: 0.75em;
            background: #000;
            padding: 10px;
            border-radius: 5px;
            max-height: 400px;
            overflow: auto;
            color: #10b981;
            white-space: pre;
            border: 1px solid #334155;
            margin-top: 10px;
            display: none;
        }

        .loader {
            display: none;
            color: #3b82f6;
            font-weight: bold;
            margin-top: 15px;
            text-align: center;
        }
        .spinner {
            display: inline-block;
            width: 12px; height: 12px;
            border: 2px solid #3b82f6;
            border-top: 2px solid transparent;
            border-radius: 50%;
            animation: spin 0.6s linear infinite;
            margin-right: 6px;
            vertical-align: middle;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>
<div class="container">
    <h1>📂 Phase 1: Folder Reader</h1>
    <p>Drag and drop a <b>folder</b> containing Tenant Excel (.xlsx, .xls) or PDF files.</p>

    <div class="drop-zone" id="dropZone">
        <div style="font-size: 3rem;">📥</div>
        <p>Drop Folder Here</p>
    </div>

    <div id="loader" class="loader">
        <span class="spinner"></span> Reading files...
    </div>

    <div class="summary" id="summary">
        <b>📊 Upload Summary</b>
        <div class="summary-grid">
            <div class="summary-card">
                <div class="num" id="totalFiles">0</div>
                <div class="label">Files Read</div>
            </div>
            <div class="summary-card">
                <div class="num" id="totalSuccess" style="color:#10b981">0</div>
                <div class="label">Successful</div>
            </div>
            <div class="summary-card">
                <div class="num" id="totalRows" style="color:#f59e0b">0</div>
                <div class="label">Total Rows</div>
            </div>
        </div>
    </div>

    <div class="file-list" id="fileList"></div>
</div>

<script>
const dropZone = document.getElementById("dropZone");
const fileList = document.getElementById("fileList");
const loader   = document.getElementById("loader");
const summary  = document.getElementById("summary");

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
    var fileCount = 0;

    for (var i = 0; i < items.length; i++) {
        var item = items[i].webkitGetAsEntry();
        if (item) {
            await traverseTree(item, formData, "");
        }
    }

    try {
        var response = await fetch("/upload", { method: "POST", body: formData });
        var result = await response.json();
        loader.style.display = "none";
        showResults(result);
    } catch (err) {
        loader.style.display = "none";
        alert("Error: " + err.message);
    }
});

function traverseTree(item, formData, path) {
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
            reader.readEntries(function(entries) {
                var promises = [];
                for (var i = 0; i < entries.length; i++) {
                    promises.push(traverseTree(entries[i], formData, path + item.name + "/"));
                }
                Promise.all(promises).then(resolve);
            });
        } else {
            resolve();
        }
    });
}

function showResults(data) {
    if (data.error) {
        alert(data.error);
        return;
    }

    var results    = data.results;
    var totalFiles = results.length;
    var totalOk    = 0;
    var totalRows  = 0;

    for (var i = 0; i < results.length; i++) {
        if (results[i].success) totalOk++;
        totalRows += results[i].total_rows || 0;
    }

    document.getElementById("totalFiles").textContent   = totalFiles;
    document.getElementById("totalSuccess").textContent  = totalOk;
    document.getElementById("totalRows").textContent     = totalRows.toLocaleString();
    summary.style.display = "block";

    for (var i = 0; i < results.length; i++) {
        var res = results[i];
        var card = document.createElement("div");
        card.className = "file-card";

        var ext = res.filename.split(".").pop().toLowerCase();
        var badgeType = ext === "pdf" ? "pdf" : ext === "xls" ? "xls" : "xlsx";
        var statusBadge = res.success ? "ok" : "fail";
        var statusText  = res.success ? "✅ OK" : "❌ FAIL";

        var previewLines = "";
        if (Array.isArray(res.data)) {
            for (var j = 0; j < res.data.length; j++) {
                var row = res.data[j];
                if (Array.isArray(row)) {
                    previewLines += row.join("  |  ") + "\n";
                } else {
                    previewLines += row + "\n";
                }
            }
        } else {
            previewLines = String(res.data);
        }

        var cardId = "preview_" + i;

        card.innerHTML =
            '<div class="file-header" onclick="togglePreview(\'' + cardId + '\')">' +
                '<div class="file-info">' +
                    '<span class="badge ' + badgeType + '">' + ext.toUpperCase() + '</span>' +
                    '<b>' + escapeHtml(res.filename) + '</b>' +
                    '<span class="badge ' + statusBadge + '">' + statusText + '</span>' +
                '</div>' +
                '<div class="row-count">' + (res.total_rows || 0).toLocaleString() + ' rows ▼</div>' +
            '</div>' +
            '<div class="preview" id="' + cardId + '">' + escapeHtml(previewLines) + '</div>';

        fileList.appendChild(card);
    }
}

function togglePreview(id) {
    var el = document.getElementById(id);
    if (el.style.display === "block") {
        el.style.display = "none";
    } else {
        el.style.display = "block";
    }
}

function escapeHtml(text) {
    var div = document.createElement("div");
    div.textContent = text;
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
        session_id = str(uuid.uuid4())[:8]
        safe_name = session_id + "_" + secure_filename(f.filename)
        filepath = UPLOAD_FOLDER / safe_name
        f.save(filepath)

        ext = filepath.suffix.lower()
        content = []
        success = True

        if ext in [".xlsx", ".xls"]:
            content = read_excel(filepath)
        elif ext == ".pdf":
            content = read_pdf(filepath)
        else:
            content = "Unsupported format"
            success = False

        if isinstance(content, str):
            success = False

        total_rows = len(content) if isinstance(content, list) else 0

        results.append({
            "filename": f.filename,
            "success": success,
            "total_rows": total_rows,
            "data": content,
        })

        if filepath.exists():
            os.remove(filepath)

    return jsonify({"results": results})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting on port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)
