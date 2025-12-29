import os
import time
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx import Document
from label_layout import (
    add_cover_label,
    add_sheet_label,
    add_workorder_label
)

os.makedirs("output", exist_ok=True)

def input_int(msg, min_val=0, max_val=None):
    while True:
        try:
            v = int(input(msg))
            if v < min_val:
                raise ValueError
            if max_val is not None and v > max_val:
                raise ValueError
            return v
        except ValueError:
            print("❌ Ingresa un número válido.")


# --------- INGRESO COVER ---------
lot_number = input("LOT number: ")

num_wo = input_int("Cantidad de Work Orders (max 4): ", 1, 4)

work_orders = []
for i in range(num_wo):
    print(f"\n--- WORK ORDER {i+1} ---")
    wo = {
        "part": input("Part #: "),
        "tag_desc": input("TAG + DESCRIPTION: "),
        "code": input("Code (barcode): "),
        "work_order": input("WO #: "),
        "remaining": input_int("Cantidad TOTAL de piezas: ", 1)
    }
    work_orders.append(wo)


# --------- DOCUMENT ---------
doc = Document()
slots = [(0,0),(0,1),(1,0),(1,1),(2,0),(2,1)]

def new_page():
    table = doc.add_table(rows=3, cols=2)
    return table, 0


table, slot_idx = new_page()

# ---------- COVER PAGE ----------
table, slot_idx = new_page()

# Slot 1 → LOT
add_cover_label(table.cell(0, 0), lot_number)

# Slots 2–5 → WO (máx 4)
slot_idx = 1
for wo in work_orders:
    if slot_idx >= 5:
        break  # slot 6 queda vacío siempre

    r, c = slots[slot_idx]

    # ✅ SOLO AÑADIDO: en cover QTY = total (remaining inicial)
    wo["qty_override"] = wo["remaining"]
    add_workorder_label(table.cell(r, c), wo, lot_number)
    del wo["qty_override"]

    slot_idx += 1

# Slot 6 → vacío (intencional)

# 🔴 FORZAR FIN DE COVER PAGE
doc.add_page_break()
table, slot_idx = new_page()

# --------- SHEETS ---------
sheet_number = 1

while any(wo["remaining"] > 0 for wo in work_orders):

    if slot_idx >= 6:
        doc.add_page_break()
        table, slot_idx = new_page()

    r,c = slots[slot_idx]
    add_sheet_label(table.cell(r,c), sheet_number)
    slot_idx += 1

    print(f"\n--- SHEET {sheet_number} ---")

    for wo in work_orders:
        if wo["remaining"] == 0:
            continue

        qty = input_int(
            f"WO {wo['work_order']} (restantes {wo['remaining']}): ",
            0,
            wo["remaining"]
        )

        for _ in range(qty):
            if slot_idx >= 6:
                doc.add_page_break()
                table, slot_idx = new_page()

            r,c = slots[slot_idx]
            wo["hide_qty"] = True
            add_workorder_label(table.cell(r,c), wo, lot_number)
            del wo["hide_qty"]

            slot_idx += 1
            wo["remaining"] -= 1

    sheet_number += 1




def add_page_x_of_y_footer(document):
    section = document.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False

    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = 1  # CENTER

    # Texto "Page "
    run = p.add_run("Page ")

    # Campo PAGE
    fld_page = OxmlElement('w:fldSimple')
    fld_page.set(qn('w:instr'), 'PAGE')
    run._r.append(fld_page)

    # Texto " of "
    p.add_run(" of ")

    # Campo NUMPAGES
    fld_pages = OxmlElement('w:fldSimple')
    fld_pages.set(qn('w:instr'), 'NUMPAGES')
    p.runs[-1]._r.append(fld_pages)



# --------- SAVE ---------
filename = f"output/WorkOrder_Labels_{int(time.time())}.docx"
add_page_x_of_y_footer(doc)
doc.save(filename)
print(f"\n✅ Documento generado:\n{filename}")
