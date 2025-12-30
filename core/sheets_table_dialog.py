from __future__ import annotations

from typing import List, Dict, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox, QLabel, QHeaderView
)

WorkOrder = Dict[str, object]
Sheet = Dict[str, object]


class SheetsTableDialog(QDialog):
    """
    Sheets editor:
    - Rows = sheets (Sheet 1, Sheet 2, ...)
    - Columns = Work Orders (PART + WO)
    - Cells = qty allocated for that WO in that sheet
    - Unlimited sheets (rows)
    - Real-time "Remaining" table updates
    - Hard rule: Column totals can NEVER exceed each WO total_qty (enforced live)
    - Save rule: Column totals MUST equal each WO total_qty
    """

    def __init__(self, parent, work_orders: List[WorkOrder], initial_sheets: Optional[List[Sheet]] = None):
        super().__init__(parent)
        self.setWindowTitle("Enter Sheets (Nesting Table)")
        self.resize(1200, 560)

        self.work_orders = work_orders
        self._result: Optional[List[Sheet]] = None

        self._updating = False
        self._last_valid = {}  # (r,c) -> int

        layout = QVBoxLayout(self)

        layout.addWidget(QLabel(
            "Enter QTY per sheet. Add as many sheets as needed.\n"
            "Rules:\n"
            "• You cannot exceed TOTAL QTY per Part/WO at any time.\n"
            "• When you click Save, column totals must match TOTAL QTY exactly."
        ))

        top = QHBoxLayout()
        layout.addLayout(top)

        # Sheets table
        self.table = QTableWidget()
        self.table.setColumnCount(len(work_orders))
        self.table.setEditTriggers(QTableWidget.AllEditTriggers)
        top.addWidget(self.table, stretch=4)

        # Column headers
        headers = []
        for wo in work_orders:
            part = str(wo.get("part", "")).strip()
            wo_num = str(wo.get("work_order", "")).strip()
            headers.append(f"PART {part}\nWO {wo_num}")
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Rows: preload if initial_sheets provided
        if initial_sheets and len(initial_sheets) > 0:
            self.table.setRowCount(len(initial_sheets))
        else:
            self.table.setRowCount(1)

        self._refresh_sheet_row_headers()

        if initial_sheets and len(initial_sheets) > 0:
            by_sheet = {}
            for sh in initial_sheets:
                sn = int(sh.get("sheet_number", 0))
                allocs = {int(i): int(q) for (i, q) in sh.get("allocations", [])}
                by_sheet[sn] = allocs

            for r in range(self.table.rowCount()):
                sn = r + 1
                allocs = by_sheet.get(sn, {})
                for c in range(self.table.columnCount()):
                    val = int(allocs.get(c, 0))
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(r, c, item)
                    self._last_valid[(r, c)] = val
        else:
            self._init_row(0)

        # Remaining table (right side)
        self.remaining_table = QTableWidget()
        self.remaining_table.setColumnCount(3)
        self.remaining_table.setRowCount(len(work_orders))
        self.remaining_table.setHorizontalHeaderLabels(["PART #", "WO #", "Remaining"])
        self.remaining_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.remaining_table.verticalHeader().setVisible(False)
        self.remaining_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        top.addWidget(self.remaining_table, stretch=2)

        self._init_remaining_table()
        self._update_remaining_table()

        # Buttons
        btn_row = QHBoxLayout()
        self.add_sheet_btn = QPushButton("Add Sheet")
        self.remove_sheet_btn = QPushButton("Remove Selected Sheet(s)")
        self.save_btn = QPushButton("Save")
        self.cancel_btn = QPushButton("Cancel")
        btn_row.addWidget(self.add_sheet_btn)
        btn_row.addWidget(self.remove_sheet_btn)
        btn_row.addStretch(1)
        btn_row.addWidget(self.save_btn)
        btn_row.addWidget(self.cancel_btn)
        layout.addLayout(btn_row)

        self.add_sheet_btn.clicked.connect(self._add_sheet)
        self.remove_sheet_btn.clicked.connect(self._remove_selected_sheets)
        self.save_btn.clicked.connect(self._save)
        self.cancel_btn.clicked.connect(self.reject)

        self.table.itemChanged.connect(self._on_item_changed)

    # ---------- Init helpers ----------
    def _refresh_sheet_row_headers(self):
        labels = [f"Sheet {r+1}" for r in range(self.table.rowCount())]
        self.table.setVerticalHeaderLabels(labels)

    def _init_row(self, r: int):
        for c in range(self.table.columnCount()):
            item = QTableWidgetItem("0")
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, c, item)
            self._last_valid[(r, c)] = 0

    def _init_remaining_table(self):
        for r, wo in enumerate(self.work_orders):
            part = str(wo.get("part", "")).strip()
            wo_num = str(wo.get("work_order", "")).strip()

            it_part = QTableWidgetItem(part)
            it_wo = QTableWidgetItem(wo_num)
            it_rem = QTableWidgetItem("")

            for it in (it_part, it_wo, it_rem):
                it.setTextAlignment(Qt.AlignCenter)

            self.remaining_table.setItem(r, 0, it_part)
            self.remaining_table.setItem(r, 1, it_wo)
            self.remaining_table.setItem(r, 2, it_rem)

    # ---------- Totals/remaining ----------
    def _expected_total(self, col: int) -> int:
        return int(self.work_orders[col].get("total_qty", 0))

    def _cell_value(self, r: int, c: int) -> int:
        item = self.table.item(r, c)
        txt = item.text().strip() if item else "0"
        if txt == "":
            return 0
        try:
            v = int(txt)
            return v if v >= 0 else 0
        except Exception:
            return 0

    def _col_sum_except_row(self, col: int, exclude_row: int) -> int:
        s = 0
        for r in range(self.table.rowCount()):
            if r == exclude_row:
                continue
            s += self._cell_value(r, col)
        return s

    def _update_remaining_table(self):
        for c in range(self.table.columnCount()):
            expected = self._expected_total(c)
            current_total = 0
            for r in range(self.table.rowCount()):
                current_total += self._cell_value(r, c)
            remaining = expected - current_total
            self.remaining_table.item(c, 2).setText(str(remaining))

            if remaining < 0:
                self.remaining_table.item(c, 2).setBackground(Qt.red)
            else:
                self.remaining_table.item(c, 2).setBackground(Qt.white)

    # ---------- Live enforcement ----------
    def _on_item_changed(self, item: QTableWidgetItem):
        if self._updating:
            return

        r = item.row()
        c = item.column()

        raw = item.text().strip()
        if raw == "":
            raw = "0"

        try:
            new_val = int(raw)
            if new_val < 0:
                raise ValueError
        except Exception:
            self._revert_cell(r, c, reason="Please enter a non-negative integer.")
            return

        expected = self._expected_total(c)
        other_sum = self._col_sum_except_row(c, r)

        if new_val + other_sum > expected:
            max_allowed = max(0, expected - other_sum)
            self._set_cell_value(r, c, max_allowed)
            QMessageBox.warning(
                self,
                "Exceeds total",
                f"You cannot exceed TOTAL QTY for this WO.\n\n"
                f"Max allowed in this cell right now: {max_allowed}"
            )
            return

        self._last_valid[(r, c)] = new_val
        self._update_remaining_table()

    def _revert_cell(self, r: int, c: int, *, reason: str):
        prev = self._last_valid.get((r, c), 0)
        self._set_cell_value(r, c, prev)
        QMessageBox.warning(self, "Invalid value", reason)

    def _set_cell_value(self, r: int, c: int, val: int):
        self._updating = True
        try:
            it = self.table.item(r, c)
            if it is None:
                it = QTableWidgetItem("0")
                it.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(r, c, it)
            it.setText(str(val))
            self._last_valid[(r, c)] = val
        finally:
            self._updating = False
        self._update_remaining_table()

    # ---------- Buttons ----------
    def _add_sheet(self):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self._init_row(r)
        self._refresh_sheet_row_headers()
        self._update_remaining_table()

    def _remove_selected_sheets(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.information(self, "Remove", "Select at least one cell in the sheet row(s) you want to remove.")
            return
        if len(rows) >= self.table.rowCount():
            QMessageBox.warning(self, "Remove", "At least one sheet row must remain.")
            return

        self._updating = True
        try:
            for rr in rows:
                self.table.removeRow(rr)

            # rebuild last_valid
            self._last_valid = {}
            for r in range(self.table.rowCount()):
                for c in range(self.table.columnCount()):
                    self._last_valid[(r, c)] = self._cell_value(r, c)
        finally:
            self._updating = False

        self._refresh_sheet_row_headers()
        self._update_remaining_table()

    # ---------- Save ----------
    def _save(self):
        totals = [0] * len(self.work_orders)

        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                if item is None:
                    item = QTableWidgetItem("0")
                    item.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(r, c, item)

                txt = item.text().strip()
                if txt == "":
                    item.setText("0")
                    txt = "0"

                try:
                    v = int(txt)
                    if v < 0:
                        raise ValueError
                except Exception:
                    QMessageBox.warning(self, "Validation", f"Row {r+1}, Column {c+1} must be a non-negative integer.")
                    return

                totals[c] += v

        for c, wo in enumerate(self.work_orders):
            expected = int(wo.get("total_qty", 0))
            if totals[c] != expected:
                part = str(wo.get("part", "")).strip()
                wo_num = str(wo.get("work_order", "")).strip()
                QMessageBox.warning(
                    self,
                    "Totals mismatch",
                    f"Column total mismatch:\nPART {part} | WO {wo_num}\n"
                    f"Entered total = {totals[c]} / Expected = {expected}\n\n"
                    "Fix the sheet allocations and try again."
                )
                return

        sheets: List[Sheet] = []
        for r in range(self.table.rowCount()):
            allocations = []
            for c in range(self.table.columnCount()):
                qty = int(self.table.item(r, c).text().strip() or "0")
                allocations.append((c, qty))
            sheets.append({"sheet_number": r + 1, "allocations": allocations})

        self._result = sheets
        self.accept()

    def result_sheets(self) -> Optional[List[Sheet]]:
        return self._result


def sheets_table_dialog(parent, work_orders: List[WorkOrder], initial_sheets: Optional[List[Sheet]] = None) -> Optional[List[Sheet]]:
    dlg = SheetsTableDialog(parent, work_orders, initial_sheets=initial_sheets)
    if dlg.exec() == QDialog.Accepted:
        return dlg.result_sheets()
    return None
