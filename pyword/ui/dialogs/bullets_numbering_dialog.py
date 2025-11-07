"""Bullets and Numbering Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QTabWidget, QWidget, QListWidget,
                             QListWidgetItem, QGroupBox, QRadioButton, QButtonGroup)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from .base_dialog import BaseDialog


class BulletsAndNumberingDialog(BaseDialog):
    """Dialog for configuring bullets and numbering."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Bullets and Numbering")
        self.list_type = "none"
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Tab widget for bullets and numbering
        self.tab_widget = QTabWidget()

        # Bullets tab
        bullets_widget = QWidget()
        bullets_layout = QVBoxLayout()

        self.bullet_list = QListWidget()
        self.bullet_list.setViewMode(QListWidget.IconMode)
        self.bullet_list.setIconSize(Qt.Size(32, 32))
        self.bullet_list.setGridSize(Qt.Size(60, 60))

        # Add bullet options
        bullet_styles = ["•", "◦", "▪", "▸", "✓", "➢", "★", "♦"]
        for bullet in bullet_styles:
            item = QListWidgetItem(bullet)
            item.setTextAlignment(Qt.AlignCenter)
            self.bullet_list.addItem(item)

        bullets_layout.addWidget(self.bullet_list)
        bullets_widget.setLayout(bullets_layout)
        self.tab_widget.addTab(bullets_widget, "Bullets")

        # Numbering tab
        numbering_widget = QWidget()
        numbering_layout = QVBoxLayout()

        self.numbering_list = QListWidget()
        self.numbering_list.setViewMode(QListWidget.IconMode)
        self.numbering_list.setIconSize(Qt.Size(32, 32))
        self.numbering_list.setGridSize(Qt.Size(100, 60))

        # Add numbering options
        numbering_styles = [
            "1. 2. 3.",
            "a. b. c.",
            "A. B. C.",
            "i. ii. iii.",
            "I. II. III.",
            "1) 2) 3)",
        ]
        for style in numbering_styles:
            item = QListWidgetItem(style)
            self.numbering_list.addItem(item)

        numbering_layout.addWidget(self.numbering_list)
        numbering_widget.setLayout(numbering_layout)
        self.tab_widget.addTab(numbering_widget, "Numbering")

        # None tab
        none_widget = QWidget()
        none_layout = QVBoxLayout()
        none_label = QLabel("Remove bullets or numbering from selected text")
        none_label.setAlignment(Qt.AlignCenter)
        none_layout.addWidget(none_label)
        none_layout.addStretch()
        none_widget.setLayout(none_layout)
        self.tab_widget.addTab(none_widget, "None")

        layout.addWidget(self.tab_widget)

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
        self.resize(400, 350)

    def get_list_type(self):
        """Get the selected list type."""
        current_tab = self.tab_widget.currentIndex()
        if current_tab == 0:  # Bullets
            return "bullet"
        elif current_tab == 1:  # Numbering
            return "numbered"
        else:  # None
            return "none"

    def get_list_style(self):
        """Get the selected list style."""
        current_tab = self.tab_widget.currentIndex()
        if current_tab == 0:  # Bullets
            item = self.bullet_list.currentItem()
            return item.text() if item else "•"
        elif current_tab == 1:  # Numbering
            item = self.numbering_list.currentItem()
            return item.text() if item else "1. 2. 3."
        return None
