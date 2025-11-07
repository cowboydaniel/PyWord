"""Borders and Shading Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QTabWidget, QWidget, QComboBox,
                             QGroupBox, QFormLayout, QSpinBox, QColorDialog,
                             QPushButton as ColorButton, QFrame)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from .base_dialog import BaseDialog


class BorderAndShadingDialog(BaseDialog):
    """Dialog for configuring borders and shading."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Borders and Shading")
        self.border_color = QColor(Qt.black)
        self.shading_color = QColor(Qt.white)
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Tab widget
        self.tab_widget = QTabWidget()

        # Borders tab
        borders_widget = self.create_borders_tab()
        self.tab_widget.addTab(borders_widget, "Borders")

        # Shading tab
        shading_widget = self.create_shading_tab()
        self.tab_widget.addTab(shading_widget, "Shading")

        layout.addWidget(self.tab_widget)

        # Preview
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()

        self.preview_frame = QFrame()
        self.preview_frame.setFrameStyle(QFrame.Box | QFrame.Plain)
        self.preview_frame.setMinimumSize(200, 100)
        self.preview_frame.setStyleSheet("background-color: white; border: 1px solid black;")
        preview_layout.addWidget(self.preview_frame)

        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)

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
        self.resize(500, 450)

    def create_borders_tab(self):
        """Create the borders configuration tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        # Border style
        style_group = QGroupBox("Style")
        style_layout = QFormLayout()

        self.border_style_combo = QComboBox()
        self.border_style_combo.addItems(["None", "Solid", "Dashed", "Dotted", "Double"])
        self.border_style_combo.currentTextChanged.connect(self.update_preview)
        style_layout.addRow("Style:", self.border_style_combo)

        self.border_width_spinbox = QSpinBox()
        self.border_width_spinbox.setMinimum(0)
        self.border_width_spinbox.setMaximum(10)
        self.border_width_spinbox.setValue(1)
        self.border_width_spinbox.setSuffix(" pt")
        self.border_width_spinbox.valueChanged.connect(self.update_preview)
        style_layout.addRow("Width:", self.border_width_spinbox)

        # Color button
        color_layout = QHBoxLayout()
        self.border_color_button = QPushButton("Choose Color...")
        self.border_color_button.clicked.connect(self.choose_border_color)
        color_layout.addWidget(self.border_color_button)
        self.border_color_display = QFrame()
        self.border_color_display.setFrameStyle(QFrame.Box)
        self.border_color_display.setFixedSize(40, 20)
        self.border_color_display.setStyleSheet("background-color: black;")
        color_layout.addWidget(self.border_color_display)
        color_layout.addStretch()
        style_layout.addRow("Color:", color_layout)

        style_group.setLayout(style_layout)
        layout.addWidget(style_group)

        # Border position
        position_group = QGroupBox("Apply To")
        position_layout = QVBoxLayout()

        self.position_combo = QComboBox()
        self.position_combo.addItems(["All", "Top", "Bottom", "Left", "Right", "Box", "None"])
        position_layout.addWidget(self.position_combo)

        position_group.setLayout(position_layout)
        layout.addWidget(position_group)

        layout.addStretch()
        widget.setLayout(layout)
        return widget

    def create_shading_tab(self):
        """Create the shading configuration tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        # Shading style
        style_group = QGroupBox("Fill")
        style_layout = QFormLayout()

        self.shading_style_combo = QComboBox()
        self.shading_style_combo.addItems(["Solid", "Clear", "Gradient"])
        self.shading_style_combo.currentTextChanged.connect(self.update_preview)
        style_layout.addRow("Style:", self.shading_style_combo)

        # Color button
        color_layout = QHBoxLayout()
        self.shading_color_button = QPushButton("Choose Color...")
        self.shading_color_button.clicked.connect(self.choose_shading_color)
        color_layout.addWidget(self.shading_color_button)
        self.shading_color_display = QFrame()
        self.shading_color_display.setFrameStyle(QFrame.Box)
        self.shading_color_display.setFixedSize(40, 20)
        self.shading_color_display.setStyleSheet("background-color: white;")
        color_layout.addWidget(self.shading_color_display)
        color_layout.addStretch()
        style_layout.addRow("Color:", color_layout)

        style_group.setLayout(style_layout)
        layout.addWidget(style_group)

        layout.addStretch()
        widget.setLayout(layout)
        return widget

    def choose_border_color(self):
        """Choose border color."""
        color = QColorDialog.getColor(self.border_color, self, "Select Border Color")
        if color.isValid():
            self.border_color = color
            self.border_color_display.setStyleSheet(f"background-color: {color.name()};")
            self.update_preview()

    def choose_shading_color(self):
        """Choose shading color."""
        color = QColorDialog.getColor(self.shading_color, self, "Select Shading Color")
        if color.isValid():
            self.shading_color = color
            self.shading_color_display.setStyleSheet(f"background-color: {color.name()};")
            self.update_preview()

    def update_preview(self):
        """Update the preview frame."""
        border_style = self.border_style_combo.currentText().lower()
        border_width = self.border_width_spinbox.value()

        style_map = {
            "solid": "solid",
            "dashed": "dashed",
            "dotted": "dotted",
            "double": "double",
            "none": "none"
        }

        border_css = style_map.get(border_style, "solid")

        if self.tab_widget.currentIndex() == 1:  # Shading tab
            bg_color = self.shading_color.name()
        else:
            bg_color = "white"

        self.preview_frame.setStyleSheet(
            f"background-color: {bg_color}; "
            f"border: {border_width}px {border_css} {self.border_color.name()};"
        )
