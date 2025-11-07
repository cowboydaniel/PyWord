"""Columns Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QSpinBox, QComboBox, QGroupBox,
                             QFormLayout, QCheckBox, QFrame)
from PySide6.QtCore import Qt
from .base_dialog import BaseDialog


class ColumnsDialog(BaseDialog):
    """Dialog for configuring document columns."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Columns")
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Presets
        presets_group = QGroupBox("Presets")
        presets_layout = QHBoxLayout()

        self.preset_combo = QComboBox()
        self.preset_combo.addItems(["One", "Two", "Three", "Left", "Right"])
        self.preset_combo.currentTextChanged.connect(self.on_preset_changed)
        presets_layout.addWidget(self.preset_combo)

        presets_group.setLayout(presets_layout)
        layout.addWidget(presets_group)

        # Number of columns
        columns_group = QGroupBox("Number of Columns")
        columns_layout = QFormLayout()

        self.columns_spinbox = QSpinBox()
        self.columns_spinbox.setMinimum(1)
        self.columns_spinbox.setMaximum(12)
        self.columns_spinbox.setValue(1)
        self.columns_spinbox.valueChanged.connect(self.update_preview)
        columns_layout.addRow("Number of columns:", self.columns_spinbox)

        columns_group.setLayout(columns_layout)
        layout.addWidget(columns_group)

        # Column properties
        properties_group = QGroupBox("Width and Spacing")
        properties_layout = QFormLayout()

        self.equal_width_checkbox = QCheckBox("Equal column width")
        self.equal_width_checkbox.setChecked(True)
        self.equal_width_checkbox.stateChanged.connect(self.on_equal_width_changed)
        properties_layout.addRow(self.equal_width_checkbox)

        self.column_width_spinbox = QSpinBox()
        self.column_width_spinbox.setMinimum(10)
        self.column_width_spinbox.setMaximum(1000)
        self.column_width_spinbox.setValue(200)
        self.column_width_spinbox.setSuffix(" pt")
        properties_layout.addRow("Width:", self.column_width_spinbox)

        self.spacing_spinbox = QSpinBox()
        self.spacing_spinbox.setMinimum(0)
        self.spacing_spinbox.setMaximum(100)
        self.spacing_spinbox.setValue(12)
        self.spacing_spinbox.setSuffix(" pt")
        self.spacing_spinbox.valueChanged.connect(self.update_preview)
        properties_layout.addRow("Spacing:", self.spacing_spinbox)

        self.line_between_checkbox = QCheckBox("Line between columns")
        properties_layout.addRow(self.line_between_checkbox)

        properties_group.setLayout(properties_layout)
        layout.addWidget(properties_group)

        # Preview
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()

        self.preview_frame = QFrame()
        self.preview_frame.setFrameStyle(QFrame.Box | QFrame.Plain)
        self.preview_frame.setMinimumSize(300, 150)
        preview_layout.addWidget(self.preview_frame)

        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)

        # Apply to
        apply_group = QGroupBox("Apply To")
        apply_layout = QVBoxLayout()

        self.apply_combo = QComboBox()
        self.apply_combo.addItems(["Whole document", "This point forward", "Selected text"])
        apply_layout.addWidget(self.apply_combo)

        apply_group.setLayout(apply_layout)
        layout.addWidget(apply_group)

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
        self.resize(400, 550)
        self.update_preview()

    def on_preset_changed(self, preset):
        """Handle preset selection."""
        preset_map = {
            "One": 1,
            "Two": 2,
            "Three": 3,
            "Left": 2,
            "Right": 2,
        }

        if preset in preset_map:
            self.columns_spinbox.setValue(preset_map[preset])

            if preset in ["Left", "Right"]:
                self.equal_width_checkbox.setChecked(False)

    def on_equal_width_changed(self, state):
        """Handle equal width checkbox change."""
        self.column_width_spinbox.setEnabled(not state)

    def update_preview(self):
        """Update the preview display."""
        # Simple visual representation of columns
        num_columns = self.columns_spinbox.value()
        spacing = self.spacing_spinbox.value()

        # Create a simple HTML preview
        preview_html = '<div style="display: flex; height: 100%; background: white;">'

        for i in range(num_columns):
            column_style = 'flex: 1; background: lightgray; margin: 5px;'
            if i < num_columns - 1:
                column_style += f' margin-right: {spacing}px;'
            preview_html += f'<div style="{column_style}"></div>'

        preview_html += '</div>'

        # Note: QFrame doesn't support HTML, this is a placeholder
        # In a real implementation, you'd use QPainter to draw the preview

    def get_column_count(self):
        """Get the number of columns."""
        return self.columns_spinbox.value()

    def get_spacing(self):
        """Get the column spacing."""
        return self.spacing_spinbox.value()

    def get_equal_width(self):
        """Check if columns should have equal width."""
        return self.equal_width_checkbox.isChecked()
