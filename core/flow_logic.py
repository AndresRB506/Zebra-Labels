from __future__ import annotations

from typing import List, Dict, Tuple, Optional
from PySide6.QtWidgets import QInputDialog, QMessageBox

WorkOrder = Dict[str, object]
Sheet = Dict[str, object]

def _gui_text(parent, title: str, label: str) -> tuple[str, bool]:
    text, ok = QInputDialog.getText(parent, title, label)
    return text.strip(), ok

def _gui_int(parent, title: str, label: str, minv: int, maxv: int, default: int = 0) -> tuple[int, bool]:
    val, ok = QInputDialog.getInt(
        parent,
        title,
        label,
        default,   # value
        minv,      # minValue
        maxv,      # maxValue
        1          # step
    )
    return int(val), ok


def collect_work_orders_via_dialogs(parent) -> Optional[List[WorkOrder]]:
    """
    Console-like sequence:
    - Ask how many WOs (1-4)
    - For each: Part, Tag+Desc, Barcode code, WO#, Total Qty
    Returns list[wo] or None if cancelled.
    """
    count, ok = _gui_int(parent, "Work Orders", "How many Work Orders? (max 4)", 1, 4, default=1)
    if not ok:
        return None

    work_orders: List[WorkOrder] = []
    for i in range(count):
        part, ok = _gui_text(parent, f"Work Order {i+1}", "Part # (required):")
        if not ok:
            return None
        if not part:
            QMessageBox.warning(parent, "Missing Part", "Part # is required.")
            return None

        tag_desc, ok = _gui_text(parent, f"Work Order {i+1}", "TAG + DESCRIPTION:")
        if not ok:
            return None

        code, ok = _gui_text(parent, f"Work Order {i+1}", "Barcode code (Code128):")
        if not ok:
            return None

        wo_num, ok = _gui_text(parent, f"Work Order {i+1}", "WO #:")
        if not ok:
            return None

        total_qty, ok = _gui_int(parent, f"Work Order {i+1}", "TOTAL QTY (entire LOT):", 1, 1_000_000, default=1)
        if not ok:
            return None

        wo: WorkOrder = {
            "part": part,
            "tag_desc": tag_desc,
            "code": code,
            "work_order": wo_num,
            "total_qty": total_qty,
        }
        work_orders.append(wo)

    return work_orders

def plan_sheets_via_dialogs(parent, work_orders: List[WorkOrder]) -> Optional[List[Sheet]]:
    """
    Console-like nesting:
    Keep creating sheets until the total allocated for each WO == total_qty.
    For each sheet, ask allocation qty for each WO (0..remaining).
    Returns list[sheet] or None if cancelled.
    """
    remaining = [int(wo["total_qty"]) for wo in work_orders]
    sheets: List[Sheet] = []
    sheet_number = 1

    while any(r > 0 for r in remaining):
        allocations: List[Tuple[int, int]] = []

        for idx, wo in enumerate(work_orders):
            if remaining[idx] <= 0:
                allocations.append((idx, 0))
                continue

            label = (
                f"WO {wo['work_order']} | PART {wo['part']}\n"
                f"Remaining: {remaining[idx]}\n\n"
                f"How many pcs go in SHEET {sheet_number}?"
            )
            qty, ok = _gui_int(parent, f"Define Sheet {sheet_number}", label, 0, remaining[idx], default=0)
            if not ok:
                return None

            allocations.append((idx, qty))

        # Update remaining
        for idx, qty in allocations:
            remaining[idx] -= qty

        sheets.append({"sheet_number": sheet_number, "allocations": allocations})
        sheet_number += 1

    return sheets

def render_summary_text(work_orders: List[WorkOrder], sheets: List[Sheet]) -> str:
    """
    Returns a console-like text summary for the GUI output panel.
    """
    totals = {i: 0 for i in range(len(work_orders))}
    lines: List[str] = []
    lines.append("================= NEST SUMMARY =================")

    for sh in sheets:
        lines.append(f"\nSHEET {sh['sheet_number']}:")
        for (i, qty) in sh["allocations"]:
            if qty > 0:
                wo = work_orders[i]
                totals[i] += qty
                lines.append(f"  - WO {wo['work_order']} | PART {wo['part']} -> {qty} pcs")

    lines.append("\nTOTALS BY WORK ORDER:")
    for i, wo in enumerate(work_orders):
        lines.append(
            f"  WO {wo['work_order']} | PART {wo['part']} -> {totals[i]} pcs "
            f"(expected {wo['total_qty']})"
        )

    lines.append("\n================================================")
    return "\n".join(lines)
