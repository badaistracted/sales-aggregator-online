# 🏬 Mall Monthly Report Automation

Automated monthly reporting system that parses tenant sales, visitor traffic, and event calendar data into a unified Excel report — with PowerPoint generation powered by AI commentary.

## 🎯 What It Does

Upload your mall's raw data files → get a polished management report.
📁 Upload Files (Sales + Traffic + Events)
↓
🔍 Auto-detect file type & parse
↓
📊 Unified Master Report (web preview)
↓
📥 Export Excel (.xlsx) — formatted, multi-tab
↓
📊 Export PowerPoint (.pptx) — charts + AI commentary

text


## 📂 Data Sources

### 1. Sales Data
- **Format:** `.xlsx`, `.xls`, or `.pdf`
- **Structures supported:**
  - **Columnar** — daily rows with Date + Sales columns
  - **Pivot** — tenants as rows, months as columns
  - **PDF Daily** — OCR-extracted daily sales reports
- Tenant names auto-detected from file content or filename

### 2. Traffic Data
- **Format:** `.xlsx` exported from Google Sheets
- **Structure:** Daily rows with date + vehicle counts (Mobil, Motor, Bus, etc.)
- Looks for `TOTAL` or `140%` column as the visitor count
- Supports Indonesian date format (`Jumat, 19 Desember 2025`)

### 3. Event Calendar
- **Format:** `.xlsx` with one tab per month
- **Structure:** Date column + location sub-columns (Main Atrium, Amphitheatre, Foodtainment, Other)
- Auto-detected by fingerprint — file must contain location keywords
- Handles:
  - Split two-row headers
  - Merged date cells (dates spanning multiple rows)
  - Sheet names with or without year (`"MEI"` or `"MEI 2026"`)
  - `PERIODE:` header rows
  - Filled or empty TIME columns

## 🏗️ Architecture
project/
├── app.py # Flask server, parsers, Excel export, routes
├── event_parser.py # Event calendar parser (multi-sheet)
├── chart_builder.py # matplotlib chart generation
├── llm_writer.py # AI commentary (OpenAI) with template fallback
├── pptx_builder.py # PowerPoint assembly (python-pptx)
├── requirements.txt
├── Dockerfile
├── README.md
├── temp_uploads/ # (auto-created) temporary file storage
└── generated_reports/ # (auto-created) exported reports

text


### Design Principles

| Principle | Implementation |
|---|---|
| **Deterministic** | Python calculates all numbers, charts, and KPIs |
| **AI for commentary only** | LLM writes executive summary & recommendations — never calculates |
| **Graceful fallback** | No OpenAI key? Template text is used automatically |
| **Simple pipeline** | Upload → Parse → Preview → Export. No database, no dashboard |

## 📊 Excel Export

Multi-tab workbook with professional formatting:

| Tab | Content |
|---|---|
| **Monthly Summary** | All tenants × all months, with totals, traffic row, sales/visitor row, events row |
| **[Tenant Name]** | Daily breakdown for each tenant: Date, Day, Day Type, Sales, Mall Traffic, Sales/Visitor, Events |
| **Events** | Full event calendar for the target month: Date, Event Name, Location, Category |

### Features
- Target month column highlighted
- Weekend rows in green
- Alternating row colors
- Grand total rows
- Freeze panes on all sheets
- IDR number formatting (`#,##0`)

## 📊 PowerPoint Export

7-slide management presentation:

| Slide | Content | Source |
|---|---|---|
| 1 | Cover | Month, timestamp, branding |
| 2 | Executive Summary | 5 KPI cards + AI paragraph |
| 3 | Monthly Sales | Bar chart + AI notes |
| 4 | Top Tenants | Horizontal bar chart + AI notes |
| 5 | Traffic & Spend | Dual-axis chart + AI notes |
| 6 | Daily Sales Pattern | Line chart with weekend shading + 7-day MA |
| 7 | Events Calendar | Event cards grouped by date |
| 8 | Recommendations | 4 AI-generated action items |

### Charts (matplotlib)
- Monthly sales bar chart (target month highlighted)
- Top 10 tenants horizontal bar
- Traffic + Sales/Visitor dual-axis (bar + line)
- Daily sales line with weekend shading, peak annotation, 7-day moving average

### AI Commentary (OpenAI)
- Model: `gpt-4o-mini` (fast, cheap)
- Temperature: `0.4` (consistent, factual)
- Structured JSON output
- Falls back to template text if no API key

## 🔍 Parser Details

### Sales Parser Pipeline
try_event_parser() → fingerprint check (Main Atrium etc.)
↓ (not event file)
try_traffic_parser() → looks for TOTAL/140% column
↓ (not traffic)
try_excel_columnar() → Date + Sales column detection
↓ (no columns found)
try_excel_pivot() → month headers as columns
↓ (no pivot found)
try_pdf_daily() → OCR/text line parsing

text


### Event Parser Pipeline
parse_event_file(filepath)
↓
fingerprint check (Main Atrium / Amphitheatre / Foodtainment)
↓ (confirmed event file)
for each sheet:
↓
detect expected month (sheet name → PERIODE header → year hint)
↓
find header row (Date column anchor)
↓
catch-all column mapping (every column between Date and Category)
↓
extract events with strict month enforcement
↓
deduplicate
↓
merge all sheets → unified event dataset

text


### Month Validation
- Every uploaded file is checked for month/year references
- Compared against the user-selected target month
- Status: ✅ Match | ⚠️ Warning | ❌ Mismatch
- Supports Indonesian month names (Januari, Februari, Maret...)

### Number Parsing
- Handles Indonesian formatting: `330.685.175` (dots as thousands separator)
- Handles Western formatting: `128,939,034`
- Handles plain numbers: `5000000`
- Handles decimals: `1,234.56`

## 🚀 Quick Start

### Local Development

```bash
# Clone and install
pip install -r requirements.txt

# Run
python app.py

# Open browser
http://localhost:5000
With AI Commentary
Bash

# Set OpenAI key (optional — works without it)
export OPENAI_API_KEY=sk-...
python app.py
Docker
Bash

# Build
docker build -t mall-report .

# Run
docker run -p 5000:5000 mall-report

# With AI
docker run -p 5000:5000 -e OPENAI_API_KEY=sk-... mall-report
Railway / Cloud Deploy
Bash

# Just push — Dockerfile handles everything
# Set OPENAI_API_KEY as environment variable (optional)
📋 Requirements
text

Python 3.11+
flask>=3.0
pandas>=2.0
openpyxl>=3.1
xlrd>=2.0
pdfplumber>=0.10
pdf2image>=1.17
pytesseract>=0.3
werkzeug>=3.0
python-pptx>=0.6.23
matplotlib>=3.9
numpy>=1.26
openai>=1.30
System Dependencies (for PDF OCR)
text

tesseract-ocr
poppler-utils
🔧 Configuration
Setting	Where	Default
Target month/year	Web UI dropdown	Current month
OpenAI API key	Environment variable OPENAI_API_KEY	None (template fallback)
Upload folder	UPLOAD_FOLDER in app.py	./temp_uploads/
Report folder	REPORTS_FOLDER in app.py	./generated_reports/
Server port	Environment variable PORT	5000
📁 Supported File Formats
Format	Sales	Traffic	Events
.xlsx	✅	✅	✅
.xls	✅	✅	✅
.pdf (digital)	✅	❌	❌
.pdf (scanned/OCR)	✅	❌	❌
🛡️ Error Handling
Files that fail to parse are listed with warnings — never crash the whole upload
Month mismatches are flagged but files are still processed
Missing traffic or events data → those sections simply omitted from exports
LLM failure → automatic template fallback, report still generates
Temporary upload files are cleaned up after processing
📝 Notes
Not a KPI dashboard — this is a report generation tool
Main output is Excel + PowerPoint — the web UI is for preview and validation
AI is optional — the system is fully functional without an OpenAI key
Calculations are deterministic — Python does all math, AI only writes text
