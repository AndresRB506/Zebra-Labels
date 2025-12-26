import os
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Pt, Inches
from barcode_utils import create_barcode_temp

FONT = "Calibri"

def clear_cell(cell):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def add_text(cell, text, size, bold=False):
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.bold = bold


# ---------- LOT SLOT ----------
def add_lot(cell, lot):
    clear_cell(cell)
    add_text(cell, "LOT", 14, True)
    add_text(cell, lot, 24, True)


# ---------- GENERIC LABEL ----------
def add_full_label(cell, wo, lot, qty_text):
    clear_cell(cell)

    add_text(cell, wo["part"], 16, True)
    add_text(cell, wo["tag"], 12)

    img = create_barcode_temp(wo["code"])
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_picture(img, width=Inches(1.45))
    os.remove(img)

    add_text(cell, f"WO {wo['work_order']}", 11)
    add_text(cell, f"LOT {lot}", 11)
    add_text(cell, qty_text, 11)

# ---------- COVER WO ----------
def add_cover_wo(cell, wo, lot):
    add_full_label(cell, wo, lot, f"QTY {wo['qty_total']}")


# ---------- SHEET SLOT ----------
def add_sheet_slot(cell, sheet):
    clear_cell(cell)
    add_text(cell, f"SHEET {sheet}", 18, True)


# ---------- SHEET WO (1 PIEZA) ----------
def add_sheet_wo(cell, wo, lot):
    add_full_label(cell, wo, lot, "QTY 1")
