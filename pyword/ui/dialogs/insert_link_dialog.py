"""Insert Link Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QLineEdit, QGroupBox, QFormLayout)
from PySide6.QtCore import Qt
from .base_dialog import BaseDialog


class InsertLinkDialog(BaseDialog):
    """Dialog for inserting a hyperlink."""

    def __init__(self, selected_text="", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Hyperlink")
        self.selected_text = selected_text
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Link details
        details_group = QGroupBox("Link Details")
        details_layout = QFormLayout()

        # Text to display
        self.text_edit = QLineEdit()
        self.text_edit.setText(self.selected_text)
        details_layout.addRow("Text to display:", self.text_edit)

        # URL
        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("https://example.com")
        details_layout.addRow("URL:", self.url_edit)

        # Tooltip
        self.tooltip_edit = QLineEdit()
        self.tooltip_edit.setPlaceholderText("Optional tooltip text")
        details_layout.addRow("Tooltip:", self.tooltip_edit)

        details_group.setLayout(details_layout)
        layout.addWidget(details_group)

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
        self.resize(400, 200)

    def get_text(self):
        """Get the link text."""
        return self.text_edit.text()

    def get_url(self):
        """Get the link URL."""
        return self.url_edit.text()

    def get_tooltip(self):
        """Get the link tooltip."""
        return self.tooltip_edit.text()


# Alias for compatibility
HyperlinkDialog = InsertLinkDialog
