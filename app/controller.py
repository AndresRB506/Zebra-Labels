from __future__ import annotations

from PySide6.QtWidgets import QMessageBox

from core.flow_logic import render_summary_text, render_work_orders_summary
from core.edit_wo_dialog import work_orders_table_dialog
from core.sheets_table_dialog import sheets_table_dialog
from core.docx_adapter import generate_doc_with_gui_color


class Controller:
    def __init__(self, ui):
        self.ui = ui

    def on_generate_full_flow(self):
        lot = self.ui.lot_input.text().strip()
        color = self.ui.color_combo.currentText().strip()

        if not lot:
            QMessageBox.warning(self.ui, "Missing LOT", "Please enter a LOT number.")
            return

        # 1) Work Orders entry (table) with summary + re-define (pre-filled)
        work_orders = None
        while True:
            work_orders = work_orders_table_dialog(
                self.ui,
                work_orders,               # <-- prefill with previous entries
                title="Enter Work Orders"
            )
            if work_orders is None:
                self._log("Cancelled while entering Work Orders.")
                return

            self.ui.output.clear()
            self._log(render_work_orders_summary(work_orders))

            redo_wos = QMessageBox.question(
                self.ui,
                "Re-define Work Orders?",
                "Do you want to re-define the Work Orders?",
                QMessageBox.Yes | QMessageBox.No
            )
            if redo_wos == QMessageBox.No:
                break

        # 2) Sheets entry (table) with summary + re-nest (pre-filled)
        sheets = None
        while True:
            sheets = sheets_table_dialog(self.ui, work_orders, initial_sheets=sheets)  # <-- prefill sheets
            if sheets is None:
                self._log("Cancelled while planning sheets.")
                return

            self.ui.output.clear()
            self._log(render_summary_text(work_orders, sheets))

            redo = QMessageBox.question(
                self.ui,
                "Re-nest sheets?",
                "Do you want to re-nest and re-enter sheet allocations?",
                QMessageBox.Yes | QMessageBox.No
            )
            if redo == QMessageBox.No:
                break

        # 3) Confirm generate
        do_gen = QMessageBox.question(
            self.ui,
            "Generate DOCX",
            "Generate/print the DOCX now?",
            QMessageBox.Yes | QMessageBox.No
        )
        if do_gen == QMessageBox.No:
            self._log("\n✅ Cancelled. No document generated.")
            return

        # 4) Generate DOCX using ORIGINAL layout (untouched)
        try:
            output_path = generate_doc_with_gui_color(
                lot_number=lot,
                work_orders=work_orders,
                sheets=sheets,
                color=color
            )
        except Exception as e:
            QMessageBox.critical(self.ui, "Error", f"Failed to generate DOCX:\n{e}")
            return

        QMessageBox.information(self.ui, "Done", f"Document generated:\n{output_path}")
        self._log(f"\n✅ Document generated:\n{output_path}")

    def _log(self, msg: str):
        self.ui.output.append(msg)
