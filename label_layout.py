from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Pt, Inches

from barcode_utils import create_barcode


# ====== FUENTES FIJAS ======
FONT_WO = ("Calibri", 30)
FONT_DESC = ("Calibri", 18)
FONT_TAG = ("Calibri", 24)
FONT_QTY = ("Calibri", 18)


def add_centered_text(cell, text, font_name, font_size, bold=False):
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run(text)
    run.bold = bold
    run.font.name = font_name
    run.font.size = Pt(font_size)


def add_cover_label(cell, lot_number):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    add_centered_text(cell, "COVER PAGE", "Arial", 12, bold=True)
    add_centered_text(cell, f"LOT # {lot_number}", "Arial", 14, bold=True)


def add_workorder_label(cell, wo):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Texto visible
    add_centered_text(
        cell,
        f"WO # {wo['work_order']}",
        FONT_WO[0],
        FONT_WO[1],
        bold=True
    )

    add_centered_text(
        cell,
        wo["description"],
        FONT_DESC[0],
        FONT_DESC[1]
    )

    add_centered_text(
        cell,
        f"TAG: {wo['tag']}",
        FONT_TAG[0],
        FONT_TAG[1]
    )

    add_centered_text(
        cell,
        f"QTY: {wo['quantity']}",
        FONT_QTY[0],
        FONT_QTY[1],
        bold=True
    )

    # BARCODE (solo barras, PART # codificado)
    barcode_value = wo["part"]
    safe_name = barcode_value.replace(" ", "_")

    barcode_file = f"output/barcode_{safe_name}"
    create_barcode(barcode_value, barcode_file)

    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(
        barcode_file + ".png",
        width=Inches(1.5)
    )
