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

# ─── CORE LOGIC: FILE READERS ────────────────────────────────

def read_excel(path):
    """Reads Excel and returns a list of rows for preview."""
    try:
        # Check if it's .xls or .xlsx
        engine = "xlrd" if path.suffix == ".xls" else "openpyxl"
        df = pd.read_excel(path, header=None, engine=engine).fillna("")
        return df.values.tolist()
    except Exception as e:
        return f"Excel Error: {str(e)}"

def read_pdf(path):
    """Reads PDF and returns text lines for preview."""
    try:
        lines = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines.extend([l.strip() for l in text.split('\n') if l.strip()])
        return lines if lines else "PDF Error: No text found."
    except Exception as e:
        return f"PDF Error: {str(e)}"

# ─── HTML/JS UI (Drag & Drop Folder) ─────────────────────────

HTML_UI = """
<!DOCTYPE html>
<html>
<head>
    <title>Phase 1: Folder Reader</title>
    <style>
        body { font-family: sans-serif; background: #0f172a; color: #e2e8f0; padding: 40px; }
        .container { max-width: 900px; margin: 0 auto; }
        .drop-zone {
            border: 3px dashed #3b82f6; border-radius: 20px;
            padding: 60px; text-align: center; cursor: pointer;
            background: rgba(59, 130, 246, 0.05); transition: 0.3s;
        }
        .drop-zone.hover { background: rgba(59, 130, 246, 0.2); border-color: #60a5fa; }
        .file-list { margin-top: 30px; }
        .file-card { 
            background: #1e293b; padding: 15px; border-radius: 10px; 
            margin-bottom: 10px; border: 1px solid #334155; 
        }
        .status { font-weight: bold; font-size: 0.8em; margin-bottom: 5px; }
        .preview { font-family: monospace; font-size: 0.75em; background: #000; 
                   padding: 10px; border-radius: 5px; max-height: 150px; overflow-y: auto; color: #10b981; }
        .loader { display: none; color: #3b82f6; font-weight: bold; margin-top: 10px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📂 Phase 1: Folder Reader</h1>
        <p>Drag and drop a <b>folder</b> containing Tenant Excel or PDF files.</p>
        
        <div class="drop-zone" id="dropZone">
            <div style="font-size: 3rem;">📥</div>
            <p>Drop Folder Here</p>
        </div>
        <div id="loader" class="loader">Processing files...</div>

        <div class="file-list" id="fileList"></div>
    </div>

    <script>
        const dropZone = document.getElementById('dropZone');
        const fileList = document.getElementById('fileList');
        const loader = document.getElementById('loader');

        dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('hover'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('hover'));

        dropZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            dropZone.classList.remove('hover');
            fileList.innerHTML = '';
            loader.style.display = 'block';

            const items = e.dataTransfer.items;
            const formData = new FormData();

            // Recursively get all files from folders
            for (let i = 0; i < items.length; i++) {
                const item = items[i].webkitGetAsEntry();
                if (item) {
                    await traverseFileTree(item, formData);
                }
            }

            // Send to Flask
            const response = await fetch('/upload', { method: 'POST', body: formData });
            const result = await response.json();
            
            loader.style.display = 'none';
            renderResults(result);
        });

        async function traverseFileTree(item, formData, path = "") {
            if (item.isFile) {
                const file = await new Promise(resolve => item.file(resolve));
                // Only accept Excel and PDF
                if (file.name.match(/\.(xlsx|xls|pdf)$/i)) {
                    formData.append('files', file, path + file.name);
                }
            } else if (item.isDirectory) {
                const dirReader = item.createReader();
                const entries = await new Promise(resolve => dirReader.readEntries(resolve));
                for (let i = 0; i < entries.length; i++) {
                    await traverseFileTree(entries[i], formData, path + item.name + "/");
                }
            }
        }

        function renderResults(data) {
            if (data.error) { alert(data.error); return; }
            data.results.forEach(res => {
                const card = document.createElement('div');
                card.className = 'file-card';
                card.innerHTML = `
                    <div class="status" style="color: ${res.success ? '#10b981' : '#ef4444'}">
                        ${res.success ? '✅ READ SUCCESS' : '❌ READ FAILED'}
                    </div>
                    <div style="margin-bottom: 8px;"><b>📄 ${res.filename}</b></div>
                    <div class="preview">${JSON.stringify(res.data, null, 2)}</div>
                `;
                fileList.appendChild(card);
            });
        }
    </script>
</body>
</html>
"""

# ─── FLASK ROUTES ───────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(HTML_UI)

@app.route('/upload', methods=['POST'])
def upload():
    if 'files' not in request.files:
        return jsonify({"error": "No files found"}), 400
    
    files = request.files.getlist('files')
    results = []
    
    # Save and Read each file
    for f in files:
        session_id = str(uuid.uuid4())[:8]
        safe_name = f"{session_id}_{secure_filename(f.filename)}"
        filepath = UPLOAD_FOLDER / safe_name
        f.save(filepath)
        
        ext = filepath.suffix.lower()
        content = []
        success = True
        
        if ext in ['.xlsx', '.xls']:
            content = read_excel(filepath)
        elif ext == '.pdf':
            content = read_pdf(filepath)
        else:
            content = "Unsupported format"
            success = False
            
        results.append({
            "filename": f.filename,
            "success": success,
            "data": content[:10] # Return first 10 rows/lines for preview
        })
        
        # Cleanup
        if filepath.exists():
            os.remove(filepath)

    return jsonify({"results": results})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
