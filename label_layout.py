import os
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Pt, Inches
from barcode_utils import create_barcode_temp


# ---------------- UTIL ----------------
def clear_cell(cell):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def add_centered_text(cell, text, size=10, bold=False):
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 🔧 ESPACIADO MÁS CÓMODO (pero aún compacto)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0

    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.bold = bold


# ---------------- COVER ----------------
def add_cover_label(cell, lot_number):
    clear_cell(cell)
    add_centered_text(cell, "LOT", size=30, bold=True)
    add_centered_text(cell, lot_number, size=30, bold=True)


# ---------------- SHEET ----------------
def add_sheet_label(cell, sheet_number):
    clear_cell(cell)
    add_centered_text(cell, "SHEET", size=24, bold=True)
    add_centered_text(cell, str(sheet_number), size=24, bold=True)


# ---------------- WORK ORDER ----------------
def add_workorder_label(cell, wo, lot_number):
    clear_cell(cell)

    # PART #
    add_centered_text(cell, wo["part"], size=18, bold=True)

    # TAG + DESCRIPTION
    add_centered_text(cell, wo["tag_desc"], size=18)

    # BARCODE
    barcode_path = create_barcode_temp(wo["code"])

    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.0

    run = p.add_run()
    run.add_picture(barcode_path, width=Inches(1.35))
    os.remove(barcode_path)

    # WO
    add_centered_text(cell, f'WO {wo["work_order"]}', size=15)

    # LOT
    add_centered_text(cell, f'LOT {lot_number}', size=15)

    # ✅ ESTA PARTE DEBE ESTAR DENTRO DE LA FUNCIÓN
    if not wo.get("hide_qty", False):
        add_centered_text(
            cell,
            f'QTY {wo.get("qty_override", 1)}',
            size=15
        )
