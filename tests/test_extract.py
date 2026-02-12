from io import BytesIO

from app.main import extract_tables_to_workbook
from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle


def _build_sample_pdf() -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    data = [["Name", "Age"], ["Alice", "30"], ["Bob", "28"]]
    table = Table(data)
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]
        )
    )
    doc.build([table])
    buffer.seek(0)
    return buffer.getvalue()


def test_extract_tables_to_workbook_finds_table():
    pdf_bytes = _build_sample_pdf()
    excel_stream, sheet_count = extract_tables_to_workbook(pdf_bytes)

    assert sheet_count == 1

    excel_stream.seek(0)
    workbook = load_workbook(filename=excel_stream)
    worksheet = workbook.active

    assert worksheet.title.startswith("page-1-table-1")

    rows = list(worksheet.iter_rows(values_only=True))
    assert rows[0] == ("Name", "Age")
    assert rows[1] == ("Alice", "30")
    assert rows[2] == ("Bob", "28")
