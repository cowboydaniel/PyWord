"""Go To Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QSpinBox, QTabWidget, QWidget,
                             QFormLayout, QLineEdit)
from PySide6.QtCore import Qt
from .base_dialog import BaseDialog


class GoToDialog(BaseDialog):
    """Dialog for navigating to a specific location in the document."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Go To")
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Tab widget for different go-to options
        self.tab_widget = QTabWidget()

        # Page tab
        page_widget = QWidget()
        page_layout = QFormLayout()

        self.page_spinbox = QSpinBox()
        self.page_spinbox.setMinimum(1)
        self.page_spinbox.setMaximum(9999)
        self.page_spinbox.setValue(1)
        page_layout.addRow("Page number:", self.page_spinbox)

        page_widget.setLayout(page_layout)
        self.tab_widget.addTab(page_widget, "Page")

        # Line tab
        line_widget = QWidget()
        line_layout = QFormLayout()

        self.line_spinbox = QSpinBox()
        self.line_spinbox.setMinimum(1)
        self.line_spinbox.setMaximum(999999)
        self.line_spinbox.setValue(1)
        line_layout.addRow("Line number:", self.line_spinbox)

        line_widget.setLayout(line_layout)
        self.tab_widget.addTab(line_widget, "Line")

        layout.addWidget(self.tab_widget)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.go_button = QPushButton("Go To")
        self.go_button.clicked.connect(self.accept)
        button_layout.addWidget(self.go_button)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.resize(300, 150)

    def get_page_number(self):
        """Get the page number to go to."""
        if self.tab_widget.currentIndex() == 0:
            return self.page_spinbox.value()
        return None

    def get_line_number(self):
        """Get the line number to go to."""
        if self.tab_widget.currentIndex() == 1:
            return self.line_spinbox.value()
        return None
