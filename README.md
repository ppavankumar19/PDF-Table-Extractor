# TablePull — PDF Table Extractor

Convert tabular data inside PDFs into an Excel workbook where every detected table gets its own sheet. The single-page UI now has a richer upload/preview experience with a 4-step flow, inline PDF preview, and built-in keyboard shortcuts.

**Live deployment (Render):** https://pdf-table-extractor-3hfa.onrender.com/

## Features
- Drag-and-drop upload at `/` with animated dropzone, file pill (name + size), inline PDF thumbnail/preview, and helper tip; 50 MB hint is shown in the UI.
- Guided 4-step UI with top stepper, status bar, and progress bar: upload PDF → analyze tables → review sheets → download Excel.
- Workbook preview: sheet tabs, sticky headers/row numbers, row/column counters, highlight-aware cells, search with hit highlighting, CSV copy, fullscreen toggle, plus keyboard shortcuts (`Ctrl/Cmd+K` focuses search, `Esc` exits fullscreen).
- Progress + status toasts and graceful warnings when no tables are found.
- Immediate Excel download after preview, auto-named `<pdf>-extracted.xlsx`, with preserved highlights and stable sheet naming.
- OCR fallback for scanned/image-only tables when Tesseract is available (see **OCR setup**).
- API: `POST /extract` streams an `.xlsx` file immediately after parsing.
- Preview API: `POST /analyze` returns JSON (table count, full rows, highlight colors) so the UI mirrors the exact workbook contents.
- Sheet naming pattern `page-{page}-table-{n}` plus `X-Table-Count` header to report how many tables were found.
- Graceful fallback sheet (`no-tables-found`) when a PDF contains no detectable tables.
- Highlight carry-over: simple colored rectangles/highlight annotations in the PDF are mapped to background fills in the corresponding Excel cells when possible.

## UI flow (`templates/index.html`)
1) **Upload:** Drop or browse for a PDF (accepts `application/pdf`); the file row shows name + size, an inline PDF preview renders, and you can clear the selection with the × button.
2) **Analyze:** Click “Analyze Tables” to call `/analyze`; the stepper, status bar, progress bar, and toasts track progress and validation (e.g., non‑PDF or empty results).
3) **Preview:** Review extracted tables with sheet tabs, row/column counters, sticky headers, search (with highlighted matches), CSV copy, and fullscreen controls. Keyboard shortcuts: `Ctrl/Cmd+K` focuses search; `Esc` exits fullscreen.
4) **Download:** The app requests `/extract`, then enables “Download Excel” with an auto-named workbook (`<pdf>-extracted.xlsx`). Toasts confirm success; files remain in memory only.

## Requirements
- Python 3.10+ (3.8+ should also work) and `pip`
- Dependencies listed in `requirements.txt`
- Tesseract binary for OCR fallback (e.g., `sudo apt-get install tesseract-ocr`) plus the `pytesseract`/`pillow` Python packages (already listed in `requirements.txt`).

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
Open http://localhost:8000 for the TablePull UI. The `/analyze` and `/extract` APIs are available on the same host.

## Run the app (after setup)
```bash
cd pdf-table-extractor
source .venv/bin/activate          # Windows: .venv\Scripts\activate
.venv/bin/uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```
- UI: open http://localhost:8000 and drop a PDF (or click the dropzone), then hit "Analyze Tables" to see the workbook preview and enable download.
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

### OCR setup (for scanned/image-only PDFs)
- Install the Tesseract binary: macOS `brew install tesseract`, Debian/Ubuntu `sudo apt-get install tesseract-ocr`, Windows: install from https://github.com/UB-Mannheim/tesseract/wiki.
- `pip install -r requirements.txt` already pulls `pytesseract` and `pillow`.
- With Tesseract present, pages that contain only raster images (no extractable text/tables) are rendered to an image, OCR’d, and parsed into rows/columns. Colored regions are approximated into Excel fills when possible.

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

Sample `/analyze` response (truncated):
```json
{
  "table_count": 1,
  "tables": [
    {
      "title": "page-1-table-1",
      "rows": [["Name", "Age"], ["Alice", "30"]],
      "highlights": [[null, null], ["FFFFF2A8", null]]
    }
  ]
}
```
- `rows` keeps the original table ordering; empty cells are returned as empty strings.
- `highlights` is a row/column-aligned matrix; values are ARGB hex strings when a colored PDF rectangle or highlight overlaps the cell, otherwise `null`.

## How it works
- Table detection uses `pdfplumber` line-based extraction with `DEFAULT_TABLE_SETTINGS` defined in `app/main.py`.
- If a page has no vector tables *and* no extractable text, the page is rendered to an image (400 DPI) and OCR’d with Tesseract; rows/columns are reconstructed and highlight colors are estimated from pixel averages.
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
.venv/bin/pytest           # OCR regression test auto-skips if Tesseract is missing
```

## Troubleshooting
- **`TesseractNotFoundError` or empty OCR results:** install the system Tesseract binary and ensure it is on your `PATH` (see *OCR setup* above). Restart the server after installing.
- **Zero tables detected in PDFs with clear lines:** raise `snap_tolerance`/`join_tolerance` in `DEFAULT_TABLE_SETTINGS` or try OCR by installing Tesseract.
- **Highlights missing:** only rectangles or highlight annotations are mapped; subtle colors may be filtered out if saturation/brightness is low. For scanned PDFs, highlight colors require the OCR path.

## Deployment notes
- For production, prefer `uvicorn app.main:app --workers <n>` behind a reverse proxy (e.g., Nginx or Caddy).
- Enforce max upload size at the proxy or ASGI server if large PDFs are expected.
- Add OCR (e.g., via `pytesseract` and `pdfplumber`'s OCR hook) if you need to support scanned/image-only PDFs; not included here.

## Limitations
- Works best on digital PDFs with clear line-drawn tables.
- OCR fallback requires the system Tesseract binary; without it, image-only PDFs still return `no-tables-found`.
- OCR highlight detection is heuristic (mean color sampling) and may miss subtle highlights.
- Processing is in-memory; very large PDFs may require streaming or chunked handling.
