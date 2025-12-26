from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Pt, Inches
import os

from barcode_utils import create_barcode_temp


# ---------- Helpers ----------
def clear_cell(cell):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def add_centered_text(cell, text, size=9, bold=False, font="Arial Narrow"):
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1

    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = "Calibri"


# ---------- COVER ----------
def add_cover_label(cell, lot_number):
    clear_cell(cell)
    add_centered_text(cell, "LOT", size=12, bold=True)
    add_centered_text(cell, lot_number, size=22, bold=True)


# ---------- WORK ORDER LABEL ----------
def add_workorder_label(cell, wo, lot_number):
    clear_cell(cell)

    # PART #
    add_centered_text(cell, wo["part"], size=24, bold=True)

    # TAG + DESCRIPTION (UN SOLO TEXTO)
    tag_text = f'{wo["tag"]} {wo["description"]}'
    add_centered_text(cell, tag_text, size=18)

    # BARCODE (Code128, solo código)
    barcode_path = create_barcode_temp(wo["code"])

    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    run = p.add_run()
    run.add_picture(barcode_path, width=Inches(1.75))

    os.remove(barcode_path)

    # WO
    add_centered_text(cell, f'WO {wo["work_order"]}', size=16)

    # LOT
    add_centered_text(cell, f'LOT {lot_number}', size=16)

    # QTY (NO bold)
    add_centered_text(cell, f'QTY {wo["quantity"]}', size=16)
