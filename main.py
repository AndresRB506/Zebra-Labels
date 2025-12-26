import os
from docx import Document
from label_layout import add_cover_label, add_workorder_label

os.makedirs("output", exist_ok=True)

work_orders = [
    {
        "part": "14X20X36 R6",
        "tag": "PLENUM",
        "description": "RAW / END CAP",
        "work_order": "146963",
        "quantity": 12
    },
    {
        "part": "18X18X36 R6",
        "tag": "PLENUM",
        "description": "DUAL END CAP",
        "work_order": "146965",
        "quantity": 24
    },
    {
        "part": "16X20X36 R6",
        "tag": "PLENUM",
        "description": "RAW / END CAP",
        "work_order": "146964",
        "quantity": 10
    },
    {
        "part": "213236 R6",
        "tag": "PLENUM",
        "description": "RAW / END CAP",
        "work_order": "146964",
        "quantity": 10
    },
    {
        "part": "1231321244 R6",
        "tag": "PLENUM",
        "description": "RAW / END CAP",
        "work_order": "146964",
        "quantity": 10
    }
]

lot_number = "08807"

doc = Document()
table = doc.add_table(rows=3, cols=2)

slots = [(0,0), (0,1), (1,0), (1,1), (2,0), (2,1)]

# Slot (0,0) → Cover
add_cover_label(table.cell(0,0), lot_number)

# Remaining slots
wo_index = 0

for row, col in slots[1:]:
    cell = table.cell(row, col)

    if wo_index < len(work_orders):
        add_workorder_label(cell, work_orders[wo_index])
        wo_index += 1   # ✅ THIS WAS MISSING
    else:
        cell.text = ""


doc.save("output/WorkOrder_Labels.docx")
