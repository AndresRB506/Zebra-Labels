import os
from docx import Document
from label_layout import add_cover_label, add_workorder_label

os.makedirs("output", exist_ok=True)

lot_number = "08807"

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
        "tag": "PLENUM",
        "description": "END CAP",
        "code": "P142036-R6-2-0BX",
        "work_order": "16428",
        "quantity": 20
    },
    {
        "part": "16X20X36 R6",
        "tag": "PLENUM",
        "description": "RAW END",
        "code": "P162036-R6-3-0BX",
        "work_order": "16429",
        "quantity": 10
    },
    {
        "part": "18X18X36 R6",
        "tag": "PLENUM",
        "description": "DUAL END",
        "code": "P181836-R6-4-0BX",
        "work_order": "16430",
        "quantity": 15
    },
    {
        "part": "20X20X36 R6",
        "tag": "PLENUM",
        "description": "RAW END",
        "code": "P202036-R6-5-0BX",
        "work_order": "16431",
        "quantity": 8
    }
]

doc = Document()

# 6 slots por página (3x2)
slots = [(0,0), (0,1), (1,0), (1,1), (2,0), (2,1)]

table = doc.add_table(rows=3, cols=2)

# Slot 0 → COVER
add_cover_label(table.cell(0, 0), lot_number)

wo_index = 0

for row, col in slots[1:]:
    if wo_index < len(work_orders):
        add_workorder_label(table.cell(row, col), work_orders[wo_index], lot_number)
        wo_index += 1

doc.save("output/WorkOrder_Labels.docx")
