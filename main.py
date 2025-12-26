from docx import Document
from label_layout import *

def ask_int(msg):
    while True:
        v = input(msg)
        if v.isdigit():
            return int(v)
        print("❌ Digite un número válido.")

def ask_qty(sheet, wo):
    while True:
        v = input(
            f"SHEET {sheet} → piezas de WO {wo['work_order']} "
            f"(restante {wo['qty_left']}): "
        )
        if not v.isdigit():
            print("❌ Debe ser un número.")
            continue

        qty = int(v)
        if qty > wo["qty_left"]:
            print(f"❌ No puede exceder {wo['qty_left']}.")
            continue

        return qty


doc = Document()
lot = input("LOT #: ")

# ---------- WORK ORDERS ----------
num_wo = ask_int("Cantidad de WO (1-4): ")
work_orders = []

for i in range(num_wo):
    print(f"\nWO {i+1}")
    wo = {
        "part": input("Part #: "),
        "tag": input("TAG + Description: "),
        "code": input("Code: "),
        "work_order": input("WO #: "),
        "qty_total": ask_int("QTY TOTAL: "),
        "qty_left": 0
    }
    wo["qty_left"] = wo["qty_total"]
    work_orders.append(wo)

# ---------- COVER PAGE ----------
table = doc.add_table(rows=3, cols=2)
slots = [(r, c) for r in range(3) for c in range(2)]

add_lot(table.cell(*slots[0]), lot)

for i, wo in enumerate(work_orders):
    add_cover_wo(table.cell(*slots[i + 1]), wo, lot)

# ---------- SLOT ENGINE ----------
current_table = None
slot_index = 0

def new_page():
    global current_table, slot_index
    doc.add_page_break()
    current_table = doc.add_table(rows=3, cols=2)
    slot_index = 0

def next_cell():
    global slot_index
    r, c = divmod(slot_index, 2)
    slot_index += 1
    return current_table.cell(r, c)

new_page()
sheet = 1

# ---------- SHEETS ----------
while any(wo["qty_left"] > 0 for wo in work_orders):

    if slot_index == 6:
        new_page()

    add_sheet_slot(next_cell(), sheet)

    for wo in work_orders:
        if wo["qty_left"] == 0:
            continue

        qty = ask_qty(sheet, wo)
        wo["qty_left"] -= qty

        for _ in range(qty):
            if slot_index == 6:
                new_page()
            add_sheet_wo(next_cell(), wo, lot)

    sheet += 1

doc.save("output/WorkOrder_Labels.docx")
print("\n✅ Documento generado correctamente.")
