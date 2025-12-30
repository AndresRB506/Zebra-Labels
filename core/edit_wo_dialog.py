from __future__ import annotations

from typing import List, Dict, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox, QLabel
)

WorkOrder = Dict[str, object]


class EditWorkOrdersDialog(QDialog):
    HEADERS = ["Part #", "TAG + DESCRIPTION", "Barcode code", "WO #", "TOTAL QTY"]

    def __init__(self, parent, work_orders: Optional[List[WorkOrder]] = None, *, title: str = "Work Orders"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(900, 420)

        if work_orders is None:
            work_orders = []

        self._result: Optional[List[WorkOrder]] = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Fill the table. Click Save to continue."))

        self.table = QTableWidget()
        self.table.setColumnCount(len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)

        # If no WOs provided, start with 1 empty row
        rows = len(work_orders) if len(work_orders) > 0 else 1
        self.table.setRowCount(rows)

        layout.addWidget(self.table)

        # Fill rows (if any)
        for r, wo in enumerate(work_orders):
            self._set_item(r, 0, str(wo.get("part", "")).strip())
            self._set_item(r, 1, str(wo.get("tag_desc", "")).strip())
            self._set_item(r, 2, str(wo.get("code", "")).strip())
            self._set_item(r, 3, str(wo.get("work_order", "")).strip())
            self._set_item(r, 4, str(int(wo.get("total_qty", 1))))

        # If started empty, prefill qty with 1
        if len(work_orders) == 0:
            self._set_item(0, 4, "1")

        btn_row = QHBoxLayout()
        self.add_btn = QPushButton("Add WO")
        self.del_btn = QPushButton("Remove Selected")
        self.save_btn = QPushButton("Save")
        self.cancel_btn = QPushButton("Cancel")
        btn_row.addWidget(self.add_btn)
        btn_row.addWidget(self.del_btn)
        btn_row.addStretch(1)
        btn_row.addWidget(self.save_btn)
        btn_row.addWidget(self.cancel_btn)
        layout.addLayout(btn_row)

        self.add_btn.clicked.connect(self._add_row)
        self.del_btn.clicked.connect(self._remove_selected)
        self.save_btn.clicked.connect(self._save)
        self.cancel_btn.clicked.connect(self.reject)

    def _set_item(self, r: int, c: int, text: str):
        item = QTableWidgetItem(text)
        if c == 4:
            item.setTextAlignment(Qt.AlignCenter)
        self.table.setItem(r, c, item)

    def _add_row(self):
        if self.table.rowCount() >= 4:
            QMessageBox.warning(self, "Limit reached", "Maximum 4 Work Orders allowed.")
            return

        r = self.table.rowCount()
        self.table.insertRow(r)
        for c in range(self.table.columnCount()):
            self._set_item(r, c, "" if c != 4 else "1")


    def _remove_selected(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.information(self, "Remove", "Select at least one cell/row to remove.")
            return
        for r in rows:
            self.table.removeRow(r)

    def _save(self):
        wos: List[WorkOrder] = []
        for r in range(self.table.rowCount()):
            part = (self.table.item(r, 0).text() if self.table.item(r, 0) else "").strip()
            tag_desc = (self.table.item(r, 1).text() if self.table.item(r, 1) else "").strip()
            code = (self.table.item(r, 2).text() if self.table.item(r, 2) else "").strip()
            wo_num = (self.table.item(r, 3).text() if self.table.item(r, 3) else "").strip()
            qty_txt = (self.table.item(r, 4).text() if self.table.item(r, 4) else "").strip()

            # Skip completely empty rows
            if not part and not tag_desc and not code and not wo_num and not qty_txt:
                continue

            if not part:
                QMessageBox.warning(self, "Validation", f"Row {r+1}: Part # is required.")
                return
            if not wo_num:
                QMessageBox.warning(self, "Validation", f"Row {r+1}: WO # is required.")
                return
            try:
                total_qty = int(qty_txt)
                if total_qty <= 0:
                    raise ValueError
            except Exception:
                QMessageBox.warning(self, "Validation", f"Row {r+1}: TOTAL QTY must be a positive integer.")
                return

            wos.append(
                {
                    "part": part,
                    "tag_desc": tag_desc,
                    "code": code,
                    "work_order": wo_num,
                    "total_qty": total_qty,
                }
            )

        # 🚨 AQUI VA EL LIMITE DE 4 WO
        if len(wos) > 4:
            QMessageBox.warning(self, "Limit reached", "Maximum 4 Work Orders allowed.")
            return

        if len(wos) == 0:
            QMessageBox.warning(self, "Validation", "You must have at least one Work Order.")
            return

        self._result = wos
        self.accept()


    def result_work_orders(self) -> Optional[List[WorkOrder]]:
        return self._result


def work_orders_table_dialog(parent, work_orders: Optional[List[WorkOrder]], *, title: str) -> Optional[List[WorkOrder]]:
    """
    Unified dialog for both initial entry and editing.
    Returns updated list if user saves, None if cancelled.
    """
    dlg = EditWorkOrdersDialog(parent, work_orders, title=title)
    if dlg.exec() == QDialog.Accepted:
        return dlg.result_work_orders()
    return None
