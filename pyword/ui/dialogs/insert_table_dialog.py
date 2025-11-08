"""Insert Table Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QSpinBox, QPushButton, QGroupBox, QFormLayout)
from PySide6.QtCore import Qt
from .base_dialog import BaseDialog


class InsertTableDialog(BaseDialog):
    """Dialog for inserting a table."""

    def __init__(self, parent=None):
        super().__init__("Insert Table", parent)
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Table size group
        size_group = QGroupBox("Table Size")
        size_layout = QFormLayout()

        # Rows
        self.rows_spinbox = QSpinBox()
        self.rows_spinbox.setMinimum(1)
        self.rows_spinbox.setMaximum(100)
        self.rows_spinbox.setValue(3)
        size_layout.addRow("Number of rows:", self.rows_spinbox)

        # Columns
        self.cols_spinbox = QSpinBox()
        self.cols_spinbox.setMinimum(1)
        self.cols_spinbox.setMaximum(50)
        self.cols_spinbox.setValue(3)
        size_layout.addRow("Number of columns:", self.cols_spinbox)

        size_group.setLayout(size_layout)
        layout.addWidget(size_group)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)
        button_layout.addWidget(self.ok_button)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.resize(300, 150)

    def get_rows(self):
        """Get the number of rows."""
        return self.rows_spinbox.value()

    def get_columns(self):
        """Get the number of columns."""
        return self.cols_spinbox.value()
