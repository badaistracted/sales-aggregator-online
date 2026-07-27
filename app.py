"""
Phase 4 & 5: Local Web UI
--------------------------
A small Flask app that wraps the Phase 1-3 backend engine:
  - drag-and-drop upload of multiple tenant Excel files
  - server-side processing (continuous timeline, ID holiday logic, workbook build)
  - JSON response used to render an in-browser dashboard (Chart.js)
  - a download endpoint for the finalized multi-tab .xlsx report

Run with:  python app.py
Then open: http://localhost:5000
"""
import io
import os
import tempfile
import uuid

from flask import Flask, jsonify, render_template, request, send_file

from backend.data_parser import load_tenants
from backend.excel_export import build_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB upload cap

# In-memory store for the most recently generated reports (single-user local app).
# token -> {"bytes": BytesIO, "filename": str}
_REPORTS: dict[str, dict] = {}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/process", methods=["POST"])
def process():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files were uploaded."}), 400

    with tempfile.TemporaryDirectory() as tmpdir:
        saved_paths = []
        for f in files:
            if not f.filename:
                continue
            if not f.filename.lower().endswith((".xlsx", ".xls")):
                continue
            path = os.path.join(tmpdir, f.filename)
            f.save(path)
            saved_paths.append(path)

        if not saved_paths:
            return jsonify({"error": "No valid .xlsx/.xls files found in the upload."}), 400

        tenants, errors = load_tenants(saved_paths)

        if not tenants:
            return jsonify({
                "error": "None of the uploaded files could be processed.",
                "file_errors": errors,
            }), 400

        wb = build_workbook(tenants)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

    token = uuid.uuid4().hex
    _REPORTS[token] = {
        "bytes": buf,
        "filename": f"K_Square_Tenant_Sales_Report_{tenants[0].year}_{tenants[0].month:02d}.xlsx",
    }

    # Recompute summary + per-day series in plain Python for the dashboard
    # (the workbook uses live formulas; this JSON payload mirrors those results).
    summary = []
    series = {}
    for t in tenants:
        weekday_sales = t.df.loc[t.df["Day Type"] == "Weekday", "Sales"]
        weekend_sales = t.df.loc[t.df["Day Type"] == "Weekend", "Sales"]
        summary.append({
            "tenant": t.tenant_name,
            "total": float(t.df["Sales"].sum()),
            "weekday_avg": float(weekday_sales.mean()) if len(weekday_sales) else 0.0,
            "weekend_avg": float(weekend_sales.mean()) if len(weekend_sales) else 0.0,
        })
        series[t.tenant_name] = {
            "dates": t.df["Date"].dt.strftime("%Y-%m-%d").tolist(),
            "sales": t.df["Sales"].tolist(),
            "day_types": t.df["Day Type"].tolist(),
        }

    month_label = f"{tenants[0].year}-{tenants[0].month:02d}"

    return jsonify({
        "month_label": month_label,
        "summary": summary,
        "series": series,
        "file_errors": errors,
        "download_token": token,
    })


@app.route("/api/download/<token>")
def download(token):
    report = _REPORTS.get(token)
    if not report:
        return jsonify({"error": "This report is no longer available. Please re-process your files."}), 404

    report["bytes"].seek(0)
    return send_file(
        report["bytes"],
        as_attachment=True,
        download_name=report["filename"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)