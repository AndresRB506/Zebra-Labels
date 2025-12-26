import os
from docx import Document
from label_layout import add_cover_label, add_workorder_label

# Carpeta de salida
os.makedirs("output", exist_ok=True)

# ---------------- DATOS ----------------
work_orders = [
    {
        "part": "12X17X36 R6",
        "tag": "PLENUM BOXED",
        "description": "RAW/END CAP",
        "code": "P121736-R6-1-0BX",
        "work_order": "16427",
        "quantity": 12
    },
    {
        "part": "14X20X36 R6",
        "tag": "PLENUM BOXED",
        "description": "RAW/END CAP",
        "code": "P142036-R6-1-0BX",
        "work_order": "16428",
        "quantity": 8
    },
    {
        "part": "18X18X36 R6",
        "tag": "PLENUM BOXED",
        "description": "DUAL END CAP",
        "code": "P181836-R6-1-0BX",
        "work_order": "16429",
        "quantity": 20
    },
    {
        "part": "16X20X36 R6",
        "tag": "PLENUM BOXED",
        "description": "RAW/END CAP",
        "code": "P162036-R6-1-0BX",
        "work_order": "16430",
        "quantity": 6
    },
    {
        "part": "20X20X36 R6",
        "tag": "PLENUM BOXED",
        "description": "RAW/END CAP",
        "code": "P202036-R6-1-0BX",
        "work_order": "16431",
        "quantity": 10
    }
]

lot_number = "08807"

# ---------------- DOCUMENTO ----------------
doc = Document()
table = doc.add_table(rows=3, cols=2)

# Slots fijos (6 por página)
slots = [(0, 0), (0, 1),
         (1, 0), (1, 1),
         (2, 0), (2, 1)]

# Slot 0 → Cover
add_cover_label(table.cell(0, 0), lot_number)

# Slots restantes → Work Orders
wo_index = 0
for row, col in slots[1:]:
    if wo_index >= len(work_orders):
        break

    add_workorder_label(
        table.cell(row, col),
        work_orders[wo_index],
        lot_number
    )
    wo_index += 1

# Guardar archivo
doc.save("output/WorkOrder_Labels.docx")
