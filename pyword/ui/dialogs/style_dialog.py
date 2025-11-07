"""Style Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QListWidget, QListWidgetItem, QGroupBox,
                             QFormLayout, QLineEdit, QFontComboBox, QSpinBox,
                             QComboBox, QCheckBox, QColorDialog, QPushButton as ColorButton)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QColor
from .base_dialog import BaseDialog


class StyleDialog(BaseDialog):
    """Dialog for managing document styles."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Styles")
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QHBoxLayout()

        # Left side - style list
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("Available Styles:"))

        self.style_list = QListWidget()
        self.populate_style_list()
        self.style_list.currentItemChanged.connect(self.on_style_selected)
        left_layout.addWidget(self.style_list)

        # Style management buttons
        style_buttons = QHBoxLayout()
        self.new_button = QPushButton("New")
        self.delete_button = QPushButton("Delete")
        self.modify_button = QPushButton("Modify")
        style_buttons.addWidget(self.new_button)
        style_buttons.addWidget(self.delete_button)
        style_buttons.addWidget(self.modify_button)
        left_layout.addLayout(style_buttons)

        layout.addLayout(left_layout, 1)

        # Right side - style properties
        right_layout = QVBoxLayout()

        # Style name
        name_group = QGroupBox("Style Name")
        name_layout = QFormLayout()
        self.name_edit = QLineEdit()
        name_layout.addRow("Name:", self.name_edit)
        name_group.setLayout(name_layout)
        right_layout.addWidget(name_group)

        # Font properties
        font_group = QGroupBox("Font")
        font_layout = QFormLayout()

        self.font_combo = QFontComboBox()
        font_layout.addRow("Font:", self.font_combo)

        self.size_spinbox = QSpinBox()
        self.size_spinbox.setMinimum(6)
        self.size_spinbox.setMaximum(72)
        self.size_spinbox.setValue(12)
        font_layout.addRow("Size:", self.size_spinbox)

        self.bold_checkbox = QCheckBox("Bold")
        self.italic_checkbox = QCheckBox("Italic")
        self.underline_checkbox = QCheckBox("Underline")
        font_layout.addRow("Style:", self.bold_checkbox)
        font_layout.addRow("", self.italic_checkbox)
        font_layout.addRow("", self.underline_checkbox)

        font_group.setLayout(font_layout)
        right_layout.addWidget(font_group)

        # Paragraph properties
        para_group = QGroupBox("Paragraph")
        para_layout = QFormLayout()

        self.alignment_combo = QComboBox()
        self.alignment_combo.addItems(["Left", "Center", "Right", "Justify"])
        para_layout.addRow("Alignment:", self.alignment_combo)

        self.line_spacing_spinbox = QSpinBox()
        self.line_spacing_spinbox.setMinimum(100)
        self.line_spacing_spinbox.setMaximum(300)
        self.line_spacing_spinbox.setValue(100)
        self.line_spacing_spinbox.setSuffix("%")
        para_layout.addRow("Line spacing:", self.line_spacing_spinbox)

        para_group.setLayout(para_layout)
        right_layout.addWidget(para_group)

        right_layout.addStretch()

        layout.addLayout(right_layout, 2)

        # Bottom buttons
        main_layout = QVBoxLayout()
        main_layout.addLayout(layout)

        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)
        button_layout.addWidget(self.ok_button)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)

        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)
        self.resize(700, 500)

    def populate_style_list(self):
        """Populate the style list with default styles."""
        styles = [
            "Normal",
            "Heading 1",
            "Heading 2",
            "Heading 3",
            "Title",
            "Subtitle",
            "Quote",
            "Code",
        ]

        for style in styles:
            self.style_list.addItem(QListWidgetItem(style))

        if self.style_list.count() > 0:
            self.style_list.setCurrentRow(0)

    def on_style_selected(self, current, previous):
        """Handle style selection."""
        if current:
            # Update the style properties based on selection
            style_name = current.text()
            self.name_edit.setText(style_name)

            # Here you would load the actual style properties
            # For now, just set some defaults based on the style name
            if "Heading" in style_name:
                self.size_spinbox.setValue(14 + (3 - int(style_name[-1])) * 2)
                self.bold_checkbox.setChecked(True)
            elif style_name == "Title":
                self.size_spinbox.setValue(18)
                self.bold_checkbox.setChecked(True)
            else:
                self.size_spinbox.setValue(12)
                self.bold_checkbox.setChecked(False)
