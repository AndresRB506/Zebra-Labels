from __future__ import annotations

from typing import List, Dict

WorkOrder = Dict[str, object]
Sheet = Dict[str, object]


def render_work_orders_summary(work_orders: List[WorkOrder]) -> str:
    """
    Console-like summary for the Work Orders table, shown in the GUI output panel.
    """
    lines: List[str] = []
    lines.append("=============== WORK ORDERS SUMMARY ===============")

    for i, wo in enumerate(work_orders, start=1):
        part = str(wo.get("part", "")).strip()
        tag = str(wo.get("tag_desc", "")).strip()
        code = str(wo.get("code", "")).strip()
        wo_num = str(wo.get("work_order", "")).strip()
        qty = int(wo.get("total_qty", 0))

        lines.append(f"\nWO #{i}")
        lines.append(f"  PART: {part}")
        lines.append(f"  WO:   {wo_num}")
        lines.append(f"  QTY:  {qty}")
        if tag:
            lines.append(f"  TAG/DESC: {tag}")
        if code:
            lines.append(f"  BARCODE:  {code}")

    lines.append("\n===================================================")
    return "\n".join(lines)


def render_summary_text(work_orders: List[WorkOrder], sheets: List[Sheet]) -> str:
    """
    Console-like nesting summary.
    sheets format:
      sheets = [{"sheet_number": 1, "allocations": [(wo_index, qty), ...]}, ...]
    """
    totals = {i: 0 for i in range(len(work_orders))}
    lines: List[str] = []
    lines.append("================= NEST SUMMARY =================")

    for sh in sheets:
        lines.append(f"\nSHEET {int(sh['sheet_number'])}:")
        for (i, qty) in sh["allocations"]:
            if int(qty) > 0:
                wo = work_orders[int(i)]
                totals[int(i)] += int(qty)
                lines.append(f"  - WO {wo['work_order']} | PART {wo['part']} -> {qty} pcs")

    lines.append("\nTOTALS BY WORK ORDER:")
    for i, wo in enumerate(work_orders):
        lines.append(
            f"  WO {wo['work_order']} | PART {wo['part']} -> {totals[i]} pcs "
            f"(expected {wo['total_qty']})"
        )

    lines.append("\n================================================")
    return "\n".join(lines)
