from fastapi import FastAPI, File, HTTPException, UploadFile, Request
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import PatternFill

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = BASE_DIR / "templates"

app = FastAPI(title="PDF Table Extractor")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))

ALLOWED_CONTENT_TYPES = {
    "application/pdf",
    "application/x-pdf",
    "application/octet-stream",  # some browsers send this for PDF uploads
}

DEFAULT_TABLE_SETTINGS: Dict[str, object] = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 3,
    "join_tolerance": 3,
}


def _ensure_pdf_bytes(pdf_bytes: bytes) -> None:
    if not pdf_bytes:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")


def _sanitize_filename(filename: Optional[str]) -> str:
    if not filename:
        return "tables"
    stem = Path(filename).stem
    safe = "".join(ch for ch in stem if ch.isalnum() or ch in {"-", "_"})
    return (safe or "tables")[:60]


def _color_to_hex(color: Optional[List[float]]) -> Optional[str]:
    """Convert a PDF RGB/gray color array to 8-char ARGB hex for openpyxl."""
    if not color:
        return None
    try:
        if len(color) == 1:
            r = g = b = int(color[0] * 255)
        else:
            r, g, b = (int(c * 255) for c in color[:3])
        return f"FF{r:02X}{g:02X}{b:02X}"
    except Exception:
        return None


def _rectangles_with_color(page) -> List[Dict[str, object]]:
    """Collect rectangles/highlight annotations that carry a fill/stroke color."""
    rects = []
    for rect in getattr(page, "rects", []):
        if rect.get("non_stroking_color") or rect.get("stroke_color"):
            rects.append(rect)

    for annot in getattr(page, "annots", []) or []:
        if annot.get("subtype", "").lower() == "highlight":
            rects.append(
                {
                    "x0": annot.get("x0"),
                    "x1": annot.get("x1"),
                    "top": annot.get("top") or annot.get("y1"),
                    "bottom": annot.get("bottom") or annot.get("y0"),
                    "non_stroking_color": annot.get("color"),
                }
            )
    return rects


def _normalize_box(obj: Dict[str, object]) -> Optional[Dict[str, float]]:
    """Return a consistent box dict or None if coordinates are missing."""
    x0 = obj.get("x0")
    x1 = obj.get("x1")
    top = obj.get("top") if obj.get("top") is not None else obj.get("y1")
    bottom = obj.get("bottom") if obj.get("bottom") is not None else obj.get("y0")
    if None in (x0, x1, top, bottom):
        return None
    return {"x0": float(x0), "x1": float(x1), "top": float(top), "bottom": float(bottom)}


def _boxes_overlap(a: Dict[str, float], b: Dict[str, float]) -> bool:
    """Basic AABB overlap check."""
    return not (
        a["x1"] <= b["x0"]
        or a["x0"] >= b["x1"]
        or a["top"] >= b["bottom"]
        or a["bottom"] <= b["top"]
    )


def _map_highlights(table, page) -> List[List[Optional[str]]]:
    """Return a matrix of ARGB colors aligned with extracted table rows.

    pdfplumber 0.11 switched `Table.cols` to `Table.columns` and now returns
    `Row`/`Column` objects with bounding boxes. This helper tolerates both the
    older numeric edges API and the newer object-based API.
    """

    data = table.extract()
    if not data:
        return []

    rows = len(data)
    cols = max(len(r) for r in data)
    highlights: List[List[Optional[str]]] = [
        [None for _ in range(cols)] for _ in range(rows)
    ]

    rects = _rectangles_with_color(page)
    if not rects:
        return highlights

    table_rows = getattr(table, "rows", []) or []
    table_cols = (
        getattr(table, "cols", None)
        if hasattr(table, "cols")
        else getattr(table, "columns", [])
    ) or []

    def _row_bounds(idx: int) -> Optional[Tuple[float, float]]:
        try:
            row_obj = table_rows[idx]
        except Exception:
            return None

        # Newer pdfplumber returns Row objects with bbox
        if hasattr(row_obj, "bbox"):
            x0, top, x1, bottom = row_obj.bbox
            return float(top), float(bottom)

        # Fallback: list/tuple shaped like a bbox
        if isinstance(row_obj, (list, tuple)) and len(row_obj) >= 4:
            return float(row_obj[1]), float(row_obj[3])

        # Fallback: numeric edge list (edges length = rows + 1)
        if isinstance(row_obj, (int, float)):
            try:
                nxt = table_rows[idx + 1]
                if isinstance(nxt, (int, float)):
                    top, bottom = float(row_obj), float(nxt)
                    return (top, bottom) if top <= bottom else (bottom, top)
            except Exception:
                return None
        return None

    def _col_bounds(idx: int) -> Optional[Tuple[float, float]]:
        try:
            col_obj = table_cols[idx]
        except Exception:
            return None

        if hasattr(col_obj, "bbox"):
            x0, top, x1, bottom = col_obj.bbox
            return float(x0), float(x1)

        if isinstance(col_obj, (list, tuple)) and len(col_obj) >= 4:
            return float(col_obj[0]), float(col_obj[2])

        if isinstance(col_obj, (int, float)):
            try:
                nxt = table_cols[idx + 1]
                if isinstance(nxt, (int, float)):
                    x0, x1 = float(col_obj), float(nxt)
                    return (x0, x1) if x0 <= x1 else (x1, x0)
            except Exception:
                return None
        return None

    for r_idx, row_cells in enumerate(data):
        bounds = _row_bounds(r_idx)
        if not bounds:
            continue
        row_top, row_bottom = bounds

        for c_idx, _ in enumerate(row_cells):
            cbounds = _col_bounds(c_idx)
            if not cbounds:
                continue
            x0, x1 = cbounds
            cell_box = {"x0": x0, "x1": x1, "top": row_top, "bottom": row_bottom}

            for rect in rects:
                rcolor = _color_to_hex(
                    rect.get("non_stroking_color") or rect.get("stroke_color")
                )
                rect_box = _normalize_box(rect)
                if rcolor and rect_box and _boxes_overlap(cell_box, rect_box):
                    highlights[r_idx][c_idx] = rcolor
                    break

    return highlights


def parse_pdf_tables(
    pdf_bytes: bytes, table_settings: Optional[Dict[str, object]] = None
) -> List[Dict[str, object]]:
    """Parse tables with optional highlight colors."""
    settings = table_settings or DEFAULT_TABLE_SETTINGS
    results: List[Dict[str, object]] = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.find_tables(table_settings=settings)
            for table_idx, table in enumerate(tables, start=1):
                data = table.extract()
                if not data:
                    continue
                highlights = _map_highlights(table, page)
                results.append(
                    {
                        "title": f"page-{page_idx}-table-{table_idx}",
                        "rows": data,
                        "highlights": highlights,
                    }
                )
    return results


def extract_tables_to_workbook(
    pdf_bytes: bytes, table_settings: Optional[Dict[str, object]] = None
) -> Tuple[BytesIO, int]:
    """Return an Excel workbook stream plus detected table count."""

    try:
        tables = parse_pdf_tables(pdf_bytes, table_settings)
        workbook = Workbook()
        sheet_count = 0

        for table in tables:
            worksheet = workbook.active if sheet_count == 0 else workbook.create_sheet()
            worksheet.title = table["title"]
            rows = table["rows"]
            highlights = table.get("highlights") or []
            for r_idx, row in enumerate(rows):
                worksheet.append(row)
                for c_idx, _ in enumerate(row):
                    try:
                        color = highlights[r_idx][c_idx]
                    except Exception:
                        color = None
                    if color:
                        worksheet.cell(row=r_idx + 1, column=c_idx + 1).fill = PatternFill(
                            start_color=color, end_color=color, fill_type="solid"
                        )
            sheet_count += 1

        if sheet_count == 0:
            worksheet = workbook.active
            worksheet.title = "no-tables-found"
            worksheet.append(["No tables were detected in this PDF."])

        output_stream = BytesIO()
        workbook.save(output_stream)
        output_stream.seek(0)
        return output_stream, sheet_count
    except Exception as exc:  # pdf could be corrupted or unsupported
        raise HTTPException(status_code=400, detail=f"Failed to parse PDF: {exc}")


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


def _is_pdf(content_type: Optional[str]) -> bool:
    if not content_type:
        return False
    lowered = content_type.lower()
    return lowered.startswith("application/pdf") or lowered in ALLOWED_CONTENT_TYPES


@app.post("/analyze")
async def analyze_tables(file: UploadFile = File(...)):
    if not _is_pdf(file.content_type):
        raise HTTPException(status_code=400, detail="Please upload a PDF file.")

    pdf_bytes = await file.read()
    _ensure_pdf_bytes(pdf_bytes)
    tables = parse_pdf_tables(pdf_bytes, table_settings=DEFAULT_TABLE_SETTINGS)
    response = {
        "table_count": len(tables),
        # Send full rows/highlights for a complete on-page preview.
        "tables": [
            {
                "title": table["title"],
                "rows": table["rows"],
                "highlights": table.get("highlights") or [],
            }
            for table in tables
        ],
    }
    return JSONResponse(response)


@app.post("/extract")
async def extract_tables(file: UploadFile = File(...)):
    if not _is_pdf(file.content_type):
        raise HTTPException(status_code=400, detail="Please upload a PDF file.")

    pdf_bytes = await file.read()
    _ensure_pdf_bytes(pdf_bytes)

    excel_stream, table_count = extract_tables_to_workbook(
        pdf_bytes, table_settings=DEFAULT_TABLE_SETTINGS
    )
    filename = _sanitize_filename(file.filename)
    disposition = f"attachment; filename={filename}-tables.xlsx"

    headers = {
        "Content-Disposition": disposition,
        "X-Table-Count": str(table_count),
    }

    return StreamingResponse(
        content=excel_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
