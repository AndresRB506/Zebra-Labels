import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Pt, Inches

from barcode_utils import create_barcode


# ---------------- UTILIDADES ----------------
def clear_cell(cell):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def add_centered_text(cell, text, size=9, bold=False, font_name="Arial"):
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1

    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font_name


# ---------------- COVER LABEL ----------------
def add_cover_label(cell, lot_number):
    clear_cell(cell)
    add_centered_text(cell, "LOT NUMBER", size=12, bold=True)
    add_centered_text(cell, lot_number, size=24, bold=True)


# ---------------- WORK ORDER LABEL ----------------
def add_workorder_label(cell, wo, lot_number):
    clear_cell(cell)

    # PART #
    add_centered_text(cell, wo["part"], size=24, bold=True)

    # 🔹 TAG + DESCRIPTION → UN SOLO TEXTO (guardado como TAG)
    tag_text = f"{wo['tag']} {wo['description']}"
    add_centered_text(cell, tag_text, size=18)

    # BARCODE (Code128, texto automático debajo)
    barcode_value = wo["code"]
    safe_name = re.sub(r"[^A-Za-z0-9_-]", "_", barcode_value)
    barcode_path = f"output/barcode_{safe_name}"

    create_barcode(barcode_value, barcode_path)

    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)

    run = p.add_run()
    run.add_picture(f"{barcode_path}.png", width=Inches(1.7))

    # WORK ORDER
    add_centered_text(cell, f"WO {wo['work_order']}", size=12)

    # LOT
    add_centered_text(cell, f"LOT {lot_number}", size=12)

    # QTY (sin negrita)
    add_centered_text(cell, f"QTY {wo['quantity']}", size=12)
