"""New Document Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QListWidget, QListWidgetItem, QPushButton, QWidget)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon, QPixmap
from .base_dialog import BaseDialog


class NewDocumentDialog(BaseDialog):
    """Dialog for creating a new document with template selection."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("New Document")
        self.selected_template = "Blank Document"
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Instructions
        label = QLabel("Select a template for your new document:")
        layout.addWidget(label)

        # Template list
        self.template_list = QListWidget()
        self.template_list.setIconSize(QSize(48, 48))
        self.template_list.setViewMode(QListWidget.ListMode)
        self.template_list.setSpacing(5)

        # Add templates
        templates = [
            ("Blank Document", "document-new"),
            ("Letter", "text-x-generic"),
            ("Resume", "x-office-document"),
            ("Report", "x-office-document"),
            ("Memo", "text-x-generic"),
        ]

        for name, icon_name in templates:
            item = QListWidgetItem(QIcon.fromTheme(icon_name), name)
            self.template_list.addItem(item)

        # Select first item by default
        self.template_list.setCurrentRow(0)
        self.template_list.itemDoubleClicked.connect(self.accept)

        layout.addWidget(self.template_list)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.ok_button = QPushButton("Create")
        self.ok_button.clicked.connect(self.accept)
        button_layout.addWidget(self.ok_button)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.resize(500, 400)

    def get_selected_template(self):
        """Get the selected template name."""
        current_item = self.template_list.currentItem()
        return current_item.text() if current_item else "Blank Document"
