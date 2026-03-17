# TablePull — PDF Table Extractor

Convert tabular data inside PDFs into an Excel workbook where every detected table gets its own sheet. The single-page UI provides a guided 4-step flow with inline PDF preview, live workbook preview, and keyboard shortcuts.

**Live deployment:** https://pdf-table-extractor-3hfa.onrender.com/

---

## Features

- **Drag-and-drop upload** with animated dropzone, file name/size pill, and inline PDF thumbnail.
- **Guided 4-step UI** — Upload → Analyze → Preview → Download — with stepper, status bar, and progress bar.
- **Live workbook preview** — sheet tabs, sticky headers, row/column counters, highlight-aware cells, search with hit highlighting, CSV copy, fullscreen toggle, and keyboard shortcuts (`Ctrl/Cmd+K` to focus search, `Esc` to exit fullscreen).
- **Excel output** — auto-named `<pdf>-tables.xlsx`, highlight fills preserved, columns auto-sized.
- **Highlight carry-over** — colored rectangles and PDF highlight annotations are mapped to Excel cell fills.
- **OCR fallback** — scanned/image-only pages are rendered and OCR'd with Tesseract when available.
- **No data stored** — all processing is in-memory; files are never written to disk.
- **50 MB upload limit** enforced on both client and server.
- **CORS-enabled** API for cross-origin use.

---

## Architecture & Data Flow

```
Browser (index.html)
│
│  1. User drops / selects a PDF file
│     • Client validates: type = application/pdf, size ≤ 50 MB
│     • PDF preview rendered via <embed> using a Blob URL
│
│  2. Click "Analyze Tables"
│     POST /analyze  (multipart/form-data, field: file)
│        │
│        ▼
│     FastAPI (app/main.py)
│        ├─ Validate content-type
│        ├─ Enforce 50 MB size limit
│        ├─ Validate PDF magic bytes (%PDF header)
│        ├─ pdfplumber → find_tables() per page
│        │    └─ If no vector tables + no text → OCR fallback (Tesseract)
│        ├─ Map highlight annotations to cell coordinates
│        └─ Return JSON: { table_count, ocr_available, tables[] }
│                           (title, rows[][], highlights[][])
│
│     Browser receives JSON:
│        • Renders sheet tabs + interactive HTML table
│        • Highlights applied as inline background colors
│
│  3. Click "Analyze Tables" (continued — Excel is built in parallel)
│     POST /extract  (same file, multipart/form-data)
│        │
│        ▼
│     FastAPI (app/main.py)
│        ├─ Same validation + parsing pipeline
│        ├─ openpyxl Workbook created
│        │    ├─ One sheet per table (page-{n}-table-{n})
│        │    ├─ PatternFill for highlight colors
│        │    ├─ Column widths auto-sized (8–60 chars)
│        │    └─ Fallback sheet "no-tables-found" if empty
│        └─ Stream binary .xlsx response
│             Headers: Content-Disposition, X-Table-Count
│
│  4. Click "Download Excel"
│     • Blob URL triggered as <a download> click
│     • File saved as <pdf-name>-extracted.xlsx
```

---

## API Reference

### `POST /analyze`

Returns a JSON preview of all detected tables — used by the UI to render the workbook preview without downloading the file.

**Request:** `multipart/form-data` with field `file` (PDF)

**Response:**
```json
{
  "table_count": 2,
  "ocr_available": false,
  "tables": [
    {
      "title": "page-1-table-1",
      "rows": [["Name", "Age"], ["Alice", "30"]],
      "highlights": [[null, null], ["FFFFF2A8", null]]
    }
  ]
}
```

- `highlights` values are 8-char ARGB hex strings (`FFRRGGBB`) when a colored annotation overlaps a cell, otherwise `null`.
- `ocr_available` indicates whether the server has Tesseract installed for scanned PDFs.

```bash
curl -X POST https://pdf-table-extractor-3hfa.onrender.com/analyze \
  -F "file=@sample.pdf" | jq .
```

---

### `POST /extract`

Streams an Excel workbook (`.xlsx`) built from all detected tables.

**Request:** `multipart/form-data` with field `file` (PDF)

**Response:** Binary `.xlsx` file

**Response headers:**
- `Content-Disposition: attachment; filename=<name>-tables.xlsx`
- `X-Table-Count: <number>`

```bash
curl -X POST https://pdf-table-extractor-3hfa.onrender.com/extract \
  -F "file=@sample.pdf" \
  -o tables.xlsx -D -
```

---

## Quickstart

```bash
git clone <your-repo-url> pdf-table-extractor
cd pdf-table-extractor
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Open http://localhost:8000 — the UI is fully self-contained (no separate static file server needed).

To stop the server: `Ctrl+C`

---

## Requirements

- Python 3.10+ (3.8+ should also work)
- Dependencies in `requirements.txt` (FastAPI, pdfplumber, openpyxl, Pillow, pytesseract)
- **Optional:** Tesseract binary for OCR on scanned PDFs

---

## OCR Setup (scanned/image-only PDFs)

Without Tesseract, pages with no extractable text return `no-tables-found`.

| Platform | Install command |
|----------|----------------|
| macOS | `brew install tesseract` |
| Debian/Ubuntu | `sudo apt-get install tesseract-ocr` |
| Windows | [UB-Mannheim installer](https://github.com/UB-Mannheim/tesseract/wiki) |

Python packages (`pytesseract`, `pillow`) are already in `requirements.txt`. Restart the server after installing Tesseract.

---

## Project Layout

```
pdf-table-extractor/
├── app/
│   ├── __init__.py
│   └── main.py          # FastAPI app — routes, PDF parsing, Excel generation
├── templates/
│   └── index.html       # Single-page UI — HTML, CSS, JavaScript
├── tests/
│   └── test_extract.py  # Regression tests (auto-skips OCR if Tesseract missing)
└── requirements.txt
```

---

## Customizing Table Detection

Tune `DEFAULT_TABLE_SETTINGS` in `app/main.py`:

```python
DEFAULT_TABLE_SETTINGS = {
    "vertical_strategy": "lines",    # or "text", "explicit"
    "horizontal_strategy": "lines",
    "snap_tolerance": 3,             # raise to merge nearby lines
    "join_tolerance": 3,
}
```

You can also pass custom settings directly to `extract_tables_to_workbook(pdf_bytes, table_settings={...})`.

---

## Testing

```bash
.venv/bin/pytest           # OCR test auto-skips if Tesseract is absent
```

---

## Deployment

**Render (current):** The app is deployed at https://pdf-table-extractor-3hfa.onrender.com/

**Self-hosted production:**
```bash
uvicorn app.main:app --workers 4 --host 0.0.0.0 --port 8000
```

Run behind a reverse proxy (Nginx, Caddy) for TLS and additional rate limiting. The 50 MB upload limit is enforced in application code; you may also set it at the proxy level.

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| `TesseractNotFoundError` | Install the system Tesseract binary and ensure it is on `PATH`. Restart the server. |
| Zero tables on a PDF with visible tables | Raise `snap_tolerance`/`join_tolerance` in `DEFAULT_TABLE_SETTINGS`. |
| Highlights missing | Only rectangles and highlight annotations are mapped. Very subtle colors (low saturation/brightness) are filtered out by design. |
| `no-tables-found` sheet in Excel | The PDF has no line-drawn tables on any page. Install Tesseract for OCR fallback on scanned PDFs. |
| Large PDF times out | The client enforces a 90 s timeout for analysis and 150 s for Excel generation. Reduce PDF size or increase server resources. |

---

## Limitations

- Works best on digital PDFs with clear line-drawn tables.
- OCR fallback is heuristic; complex layouts may produce imperfect columns.
- OCR highlight detection uses mean pixel color sampling and may miss subtle or overlapping colors.
- All processing is in-memory; very large PDFs (near 50 MB) with many pages will use proportionally more RAM.
