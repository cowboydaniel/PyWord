from enum import Enum, auto
from typing import List, Optional, Dict, Any, Tuple
from dataclasses import dataclass, field
from PySide6.QtCore import Qt, QObject, Signal
from PySide6.QtGui import QTextDocument, QTextFrame, QTextFrameFormat, QTextCursor, QTextBlockFormat, QTextCharFormat, QTextFormat
from PySide6.QtWidgets import QTextEdit, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QSpinBox, QComboBox, QDialogButtonBox, QFormLayout, QDoubleSpinBox

class ColumnLayout(Enum):
    ONE_COLUMN = 1
    TWO_COLUMNS = 2
    THREE_COLUMNS = 3
    LEFT = -1  # Two columns, narrow left
    RIGHT = -2  # Two columns, narrow right
    
    def get_column_widths(self) -> List[float]:
        """Get relative column widths for this layout."""
        if self == ColumnLayout.ONE_COLUMN:
            return [1.0]
        elif self == ColumnLayout.TWO_COLUMNS:
            return [0.5, 0.5]
        elif self == ColumnLayout.THREE_COLUMNS:
            return [0.33, 0.34, 0.33]
        elif self == ColumnLayout.LEFT:
            return [0.33, 0.67]
        elif self == ColumnLayout.RIGHT:
            return [0.67, 0.33]
        return [1.0]

@dataclass
class ColumnSettings:
    """Settings for a multi-column layout."""
    layout: ColumnLayout = ColumnLayout.ONE_COLUMN
    spacing: float = 12.0  # In points
    line_between: bool = False
    equal_width: bool = True
    custom_widths: List[float] = field(default_factory=list)
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'layout': self.layout.name,
            'spacing': self.spacing,
            'line_between': self.line_between,
            'equal_width': self.equal_width,
            'custom_widths': self.custom_widths
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ColumnSettings':
        try:
            layout = ColumnLayout[data.get('layout', 'ONE_COLUMN')]
        except KeyError:
            layout = ColumnLayout.ONE_COLUMN
            
        return cls(
            layout=layout,
            spacing=data.get('spacing', 12.0),
            line_between=data.get('line_between', False),
            equal_width=data.get('equal_width', True),
            custom_widths=data.get('custom_widths', [])
        )

class ColumnManager(QObject):
    """Manages column layout for a QTextDocument."""
    columnLayoutChanged = Signal()
    
    def __init__(self, document: QTextDocument, parent=None):
        super().__init__(parent)
        self.document = document
        self.settings = ColumnSettings()
        self._current_section = 0
        self._sections: List[ColumnSettings] = [self.settings]
        
        # Connect to document's layout changes
        self.document.documentLayout().documentSizeChanged.connect(self._update_columns)
    
    def set_layout(self, layout: ColumnLayout):
        """Set the column layout for the current section."""
        if self.settings.layout != layout:
            self.settings.layout = layout
            self._update_columns()
            self.columnLayoutChanged.emit()
    
    def set_spacing(self, spacing: float):
        """Set the spacing between columns in points."""
        if self.settings.spacing != spacing:
            self.settings.spacing = spacing
            self._update_columns()
            self.columnLayoutChanged.emit()
    
    def set_line_between(self, enabled: bool):
        """Enable or disable line between columns."""
        if self.settings.line_between != enabled:
            self.settings.line_between = enabled
            self._update_columns()
            self.columnLayoutChanged.emit()
    
    def set_equal_width(self, enabled: bool):
        """Set whether columns should have equal width."""
        if self.settings.equal_width != enabled:
            self.settings.equal_width = enabled
            self._update_columns()
            self.columnLayoutChanged.emit()
    
    def set_custom_widths(self, widths: List[float]):
        """Set custom column widths (as fractions of total width)."""
        if sum(widths) != 1.0 or any(w <= 0 for w in widths):
            raise ValueError("Column widths must be positive and sum to 1.0")
            
        self.settings.custom_widths = widths
        self.settings.equal_width = False
        self._update_columns()
        self.columnLayoutChanged.emit()
    
    def insert_column_break(self, cursor: QTextCursor):
        """Insert a column break at the current cursor position."""
        # Save the current format
        block_format = cursor.blockFormat()
        char_format = cursor.charFormat()
        
        # Insert a frame break
        cursor.insertBlock()
        
        # Create a new frame format for the column break
        frame_format = QTextFrameFormat()
        frame_format.setPageBreakPolicy(QTextFormat.PageBreak_AlwaysAfter)
        
        # Insert the frame break
        cursor.insertFrame(frame_format)
        
        # Insert a new paragraph after the break
        cursor.insertBlock(block_format, char_format)
        
        # Update columns to reflect the break
        self._update_columns()
    
    def _update_columns(self):
        """Update the document's column layout."""
        if self.settings.layout == ColumnLayout.ONE_COLUMN:
            # Reset to single column
            root_frame = self.document.rootFrame()
            frame_format = root_frame.frameFormat()
            frame_format.setColumns(1)
            frame_format.setColumnWidth(0, 0)  # 0 means use all available width
            root_frame.setFrameFormat(frame_format)
            return
        
        # Get the number of columns and their widths
        num_columns = abs(self.settings.layout.value)
        
        # Get column widths
        if self.settings.equal_width or not self.settings.custom_widths:
            widths = self.settings.layout.get_column_widths()
        else:
            widths = self.settings.custom_widths
            num_columns = len(widths)
        
        # Set up the root frame format
        root_frame = self.document.rootFrame()
        frame_format = root_frame.frameFormat()
        
        # Set column properties
        frame_format.setColumnCount(num_columns)
        frame_format.setColumnWidths([int(w * 5000) for w in widths])  # Scale for better precision
        frame_format.setColumnSpacing(self.settings.spacing)
        
        # Add line between columns if enabled
        if self.settings.line_between and num_columns > 1:
            frame_format.setProperty(QTextFormat.Property.FrameBorder, 1)
            frame_format.setProperty(QTextFormat.Property.FrameBorderBrush, Qt.GlobalColor.gray)
            frame_format.setProperty(QTextFormat.Property.FrameBorderStyle, Qt.PenStyle.SolidLine)
        else:
            frame_format.clearProperty(QTextFormat.Property.FrameBorder)
        
        # Apply the format
        root_frame.setFrameFormat(frame_format)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert column settings to a dictionary for serialization."""
        return self.settings.to_dict()
    
    def from_dict(self, data: Dict[str, Any]):
        """Load column settings from a dictionary."""
        self.settings = ColumnSettings.from_dict(data)
        self._update_columns()
        self.columnLayoutChanged.emit()

class ColumnDialog(QDialog):
    """Dialog for configuring column layout."""
    def __init__(self, column_manager: 'ColumnManager', parent=None):
        super().__init__(parent)
        self.column_manager = column_manager
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("Columns")
        self.setMinimumSize(400, 250)
        
        layout = QVBoxLayout()
        
        # Preset layouts
        presets_layout = QHBoxLayout()
        presets_layout.addWidget(QLabel("Presets:"))
        
        # Create preset buttons
        self.preset_buttons = {}
        for i, (name, col_layout) in enumerate([
            ("One", ColumnLayout.ONE_COLUMN),
            ("Two", ColumnLayout.TWO_COLUMNS),
            ("Three", ColumnLayout.THREE_COLUMNS),
            ("Left", ColumnLayout.LEFT),
            ("Right", ColumnLayout.RIGHT)
        ]):
            btn = QPushButton(name)
            btn.setCheckable(True)
            btn.setChecked(col_layout == self.column_manager.settings.layout)
            btn.clicked.connect(lambda _, l=col_layout: self._on_preset_selected(l))
            presets_layout.addWidget(btn)
            self.preset_buttons[col_layout] = btn
        
        layout.addLayout(presets_layout)
        
        # Spacing
        spacing_layout = QHBoxLayout()
        spacing_layout.addWidget(QLabel("Spacing:"))
        
        self.spacing_spin = QDoubleSpinBox()
        self.spacing_spin.setRange(0, 100)
        self.spacing_spin.setValue(self.column_manager.settings.spacing)
        self.spacing_spin.setSuffix(" pt")
        self.spacing_spin.valueChanged.connect(self._on_spacing_changed)
        spacing_layout.addWidget(self.spacing_spin)
        
        layout.addLayout(spacing_layout)
        
        # Line between
        self.line_checkbox = QCheckBox("Line between columns")
        self.line_checkbox.setChecked(self.column_manager.settings.line_between)
        self.line_checkbox.toggled.connect(self._on_line_toggled)
        layout.addWidget(self.line_checkbox)
        
        # Equal width columns (only for custom layouts)
        self.equal_width_checkbox = QCheckBox("Equal column width")
        self.equal_width_checkbox.setChecked(self.column_manager.settings.equal_width)
        self.equal_width_checkbox.toggled.connect(self._on_equal_width_toggled)
        layout.addWidget(self.equal_width_checkbox)
        
        # Preview
        preview_label = QLabel("<b>Preview:</b> " + self._get_preview_text())
        layout.addWidget(preview_label)
        
        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel |
            QDialogButtonBox.StandardButton.Apply
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        button_box.button(QDialogButtonBox.StandardButton.Apply).clicked.connect(self._apply_changes)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def _get_preview_text(self) -> str:
        """Get a preview text for the current column layout."""
        layout = self.column_manager.settings.layout
        if layout == ColumnLayout.ONE_COLUMN:
            return "Single column"
        elif layout == ColumnLayout.TWO_COLUMNS:
            return "Two equal columns"
        elif layout == ColumnLayout.THREE_COLUMNS:
            return "Three equal columns"
        elif layout == ColumnLayout.LEFT:
            return "Left narrow, right wide"
        elif layout == ColumnLayout.RIGHT:
            return "Left wide, right narrow"
        return "Custom layout"
    
    def _on_preset_selected(self, layout: ColumnLayout):
        """Handle preset column layout selection."""
        # Update button states
        for btn_layout, btn in self.preset_buttons.items():
            btn.setChecked(btn_layout == layout)
        
        # Update the column manager
        self.column_manager.set_layout(layout)
    
    def _on_spacing_changed(self, value: float):
        """Handle spacing change."""
        self.column_manager.set_spacing(value)
    
    def _on_line_toggled(self, checked: bool):
        """Handle line between columns toggle."""
        self.column_manager.set_line_between(checked)
    
    def _on_equal_width_toggled(self, checked: bool):
        """Handle equal width columns toggle."""
        self.column_manager.set_equal_width(checked)
    
    def _apply_changes(self):
        """Apply changes without closing the dialog."""
        # All changes are applied immediately through signals
        pass
    
    def accept(self):
        """Handle dialog acceptance."""
        self._apply_changes()
        super().accept()
