from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QTextEdit
)

from .controller import Controller

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Zebra Labels")
        self.resize(950, 700)

        self.ctrl = Controller(self)

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        # LOT row
        lot_row = QHBoxLayout()
        lot_row.addWidget(QLabel("LOT #:"))
        self.lot_input = QLineEdit()
        self.lot_input.setPlaceholderText("e.g. 08800")
        lot_row.addWidget(self.lot_input)
        layout.addLayout(lot_row)

        # Color row
        color_row = QHBoxLayout()
        color_row.addWidget(QLabel("Color:"))
        self.color_combo = QComboBox()
        self.color_combo.addItems(["WHITE", "ORANGE", "GREEN", "YELLOW"])
        color_row.addWidget(self.color_combo)
        layout.addLayout(color_row)

        # Buttons row
        btn_row = QHBoxLayout()
        self.generate_btn = QPushButton("Generate Word (Full Flow)")
        btn_row.addWidget(self.generate_btn)
        layout.addLayout(btn_row)

        # Output / log
        self.output = QTextEdit()
        self.output.setReadOnly(True)
        layout.addWidget(self.output)

        # Wire events
        self.generate_btn.clicked.connect(self.ctrl.on_generate_full_flow)
