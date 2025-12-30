from __future__ import annotations

from typing import List, Dict, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox, QLabel
)

WorkOrder = Dict[str, object]


class EditWorkOrdersDialog(QDialog):
    """
    Simple editor to modify existing Work Orders before nesting.
    UI-only change. DOCX layout logic remains untouched.
    """
    HEADERS = ["Part #", "TAG + DESCRIPTION", "Barcode code", "WO #", "TOTAL QTY"]

    def __init__(self, parent, work_orders: List[WorkOrder]):
        super().__init__(parent)
        self.setWindowTitle("Edit Work Orders")
        self.resize(900, 420)

        self._original = work_orders
        self._result: Optional[List[WorkOrder]] = None

        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("Edit the fields below. Click Save to apply changes."))

        self.table = QTableWidget()
        self.table.setColumnCount(len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setRowCount(len(work_orders))
        self.table.setEditTriggers(QTableWidget.AllEditTriggers)
        layout.addWidget(self.table)

        # Fill table
        for r, wo in enumerate(work_orders):
            self._set_item(r, 0, str(wo.get("part", "")).strip())
            self._set_item(r, 1, str(wo.get("tag_desc", "")).strip())
            self._set_item(r, 2, str(wo.get("code", "")).strip())
            self._set_item(r, 3, str(wo.get("work_order", "")).strip())
            self._set_item(r, 4, str(int(wo.get("total_qty", 0))))

        # Buttons
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

        if len(wos) == 0:
            QMessageBox.warning(self, "Validation", "You must have at least one Work Order.")
            return

        self._result = wos
        self.accept()

    def result_work_orders(self) -> Optional[List[WorkOrder]]:
        return self._result


def edit_work_orders_dialog(parent, work_orders: List[WorkOrder]) -> Optional[List[WorkOrder]]:
    """
    Returns updated list if user saves, None if cancelled.
    """
    dlg = EditWorkOrdersDialog(parent, work_orders)
    if dlg.exec() == QDialog.Accepted:
        return dlg.result_work_orders()
    return None
