# PDF Table Extractor

Convert tabular data inside PDFs into an Excel workbook where every detected table gets its own sheet. Includes a bold, single-page UI with full Excel-style preview (all rows, highlight colors) and REST APIs.

## Features
- Drag-and-drop upload at `/`; inline CSS/JS (no build step).
- Full workbook preview before download: sheet tabs, sticky headers/row numbers, highlight-aware cells, and PDF preview side-by-side.
- Immediate Excel download after preview; warns gracefully when no tables are found.
- API: `POST /extract` streams an `.xlsx` file immediately after parsing.
- Preview API: `POST /analyze` returns JSON (table count, full rows, highlight colors) so the UI mirrors the exact workbook contents.
- Sheet naming pattern `page-{page}-table-{n}` plus `X-Table-Count` header to report how many tables were found.
- Graceful fallback sheet (`no-tables-found`) when a PDF contains no detectable tables.
- Highlight carry-over: simple colored rectangles/highlight annotations in the PDF are mapped to background fills in the corresponding Excel cells when possible.

## Requirements
- Python 3.10+ (3.8+ should also work) and `pip`
- Dependencies listed in `requirements.txt`

## Quickstart (copy/paste)
```bash
git clone <your-repo-url> pdf-table-extractor
cd pdf-table-extractor
python3 -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
.venv/bin/pytest            # optional: run regression test
.venv/bin/uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```
Open http://localhost:8000 for the drag-and-drop UI. The `/analyze` and `/extract` APIs are available on the same host.

## Run the app (after setup)
```bash
cd pdf-table-extractor
source .venv/bin/activate          # Windows: .venv\Scripts\activate
.venv/bin/uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```
- UI: open http://localhost:8000 and drop a PDF.
- Preview JSON: `curl -X POST http://localhost:8000/analyze -F "file=@sample.pdf"`
- Excel download: `curl -X POST http://localhost:8000/extract -F "file=@sample.pdf" -o tables.xlsx`
- Stop the server: Ctrl+C (or `pkill -f "uvicorn app.main:app"` if running in the background).

## Installation & local run
```bash
git clone <your-repo-url> pdf-table-extractor
cd pdf-table-extractor
python3 -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
.venv/bin/uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```
Then open http://localhost:8000 to use the upload form (the page is self-contained; no static files needed).

## API usage
- **Endpoint:** `POST /extract` (download Excel)
- **Preview endpoint:** `POST /analyze` (JSON with table_count, all rows, highlight colors)
- **Content-Type:** `multipart/form-data`
- **Field:** `file` (PDF to process)
- **Accepted upload types:** content types starting with `application/pdf` or common aliases (`application/x-pdf`, `application/octet-stream`)

Example with `curl`:
```bash
curl -X POST http://localhost:8000/extract \
  -F "file=@sample.pdf" \
  -o tables.xlsx -D headers.txt
```
Response headers:
- `Content-Disposition`: `attachment; filename=<original-name>-tables.xlsx`
- `X-Table-Count`: number of tables detected across all pages

The body is an Excel workbook (`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`).

Preview-only example:
```bash
curl -X POST http://localhost:8000/analyze \
  -F "file=@sample.pdf" | jq .
```
Returns JSON with `table_count`, full `rows`, and per-cell ARGB `highlights` so you can mirror the Excel output in your own UI.

## How it works
- Table detection uses `pdfplumber` line-based extraction with `DEFAULT_TABLE_SETTINGS` defined in `app/main.py`.
- Each table is appended to an `openpyxl` workbook; sheets are created on the fly per table.
- All processing happens in memory (no temp files written to disk) and is returned as a streaming response.

## Customizing table detection
Tune `DEFAULT_TABLE_SETTINGS` in `app/main.py` to change how tables are found:
- `vertical_strategy` / `horizontal_strategy`: extraction strategy; currently `"lines"` for line-drawn tables.
- `snap_tolerance`, `join_tolerance`: how aggressively to merge nearby lines.

If you call `extract_tables_to_workbook(pdf_bytes, table_settings=...)` directly, pass your own settings dict to experiment without changing the defaults.

## Project layout
- `app/main.py`: FastAPI app, PDF parsing, table/highlight mapping, routes (`/`, `/analyze`, `/extract`)
- `templates/index.html`: Single-page UI with inline styles, PDF/table previews, and client logic for analyze+download flow
- `tests/test_extract.py`: regression test that builds a sample PDF and validates extraction (update for new behavior as needed)
- `requirements.txt`: runtime and test dependencies

## Testing
```bash
.venv/bin/pytest
```

## Deployment notes
- For production, prefer `uvicorn app.main:app --workers <n>` behind a reverse proxy (e.g., Nginx or Caddy).
- Enforce max upload size at the proxy or ASGI server if large PDFs are expected.
- Add OCR (e.g., via `pytesseract` and `pdfplumber`'s OCR hook) if you need to support scanned/image-only PDFs; not included here.

## Limitations
- Works best on digital PDFs with clear line-drawn tables.
- Image-only or poorly scanned PDFs will not yield tables without OCR.
- Processing is in-memory; very large PDFs may require streaming or chunked handling.
