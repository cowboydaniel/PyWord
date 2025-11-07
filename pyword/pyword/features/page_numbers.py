from enum import Enum, auto
from typing import Optional, Dict, Any, List, Tuple
from PySide6.QtCore import Qt, QObject, Signal, QRectF, QPointF
from PySide6.QtGui import QTextDocument, QTextFrame, QTextFrameFormat, QTextCharFormat, QTextCursor, QTextBlockFormat, QTextFormat, QTextBlock
from PySide6.QtWidgets import QTextEdit, QComboBox, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QDialogButtonBox, QFormLayout, QSpinBox, QCheckBox

class PageNumberPosition(Enum):
    TOP_LEFT = auto()
    TOP_CENTER = auto()
    TOP_RIGHT = auto()
    BOTTOM_LEFT = auto()
    BOTTOM_CENTER = auto()
    BOTTOM_RIGHT = auto()
    PAGE_X_OF_Y = auto()  # Special format: "Page X of Y"

class PageNumberFormat(Enum):
    NUMERIC = "1, 2, 3, ..."
    LOWER_ROMAN = "i, ii, iii, ..."
    UPPER_ROMAN = "I, II, III, ..."
    LOWER_ALPHA = "a, b, c, ..."
    UPPER_ALPHA = "A, B, C, ..."

class PageNumberSettings:
    def __init__(self):
        self.enabled = False
        self.position = PageNumberPosition.BOTTOM_CENTER
        self.format = PageNumberFormat.NUMERIC
        self.start_from = 1
        self.show_on_first_page = True
        self.alignment = Qt.AlignmentFlag.AlignCenter
        
        # Format strings for different positions
        self.format_strings = {
            PageNumberPosition.TOP_LEFT: "{number}",
            PageNumberPosition.TOP_CENTER: "{number}",
            PageNumberPosition.TOP_RIGHT: "{number}",
            PageNumberPosition.BOTTOM_LEFT: "{number}",
            PageNumberPosition.BOTTOM_CENTER: "{number}",
            PageNumberPosition.BOTTOM_RIGHT: "{number}",
            PageNumberPosition.PAGE_X_OF_Y: "Page {number} of {total}"
        }
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'enabled': self.enabled,
            'position': self.position.name,
            'format': self.format.name,
            'start_from': self.start_from,
            'show_on_first_page': self.show_on_first_page,
            'format_strings': {k.name: v for k, v in self.format_strings.items()}
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'PageNumberSettings':
        settings = cls()
        settings.enabled = data.get('enabled', False)
        settings.position = PageNumberPosition[data.get('position', 'BOTTOM_CENTER')]
        settings.format = PageNumberFormat[data.get('format', 'NUMERIC')]
        settings.start_from = data.get('start_from', 1)
        settings.show_on_first_page = data.get('show_on_first_page', True)
        
        format_strings = data.get('format_strings', {})
        for pos_name, fmt_str in format_strings.items():
            try:
                pos = PageNumberPosition[pos_name]
                settings.format_strings[pos] = fmt_str
            except KeyError:
                continue
        
        return settings

class PageNumberManager(QObject):
    """Manages page numbers in a QTextDocument."""
    pageNumbersChanged = Signal()
    
    def __init__(self, document: QTextDocument, parent=None):
        super().__init__(parent)
        self.document = document
        self.settings = PageNumberSettings()
        self._total_pages = 1
        self._current_page = 1
        
        # Connect to document's layout changes to update page numbers
        self.document.contentsChanged.connect(self._on_contents_changed)
        self.document.documentLayout().documentSizeChanged.connect(self._update_page_numbers)
    
    def set_enabled(self, enabled: bool):
        """Enable or disable page numbers."""
        if self.settings.enabled != enabled:
            self.settings.enabled = enabled
            self._update_page_numbers()
            self.pageNumbersChanged.emit()
    
    def set_position(self, position: PageNumberPosition):
        """Set the position of page numbers."""
        if self.settings.position != position:
            self.settings.position = position
            self._update_page_numbers()
            self.pageNumbersChanged.emit()
    
    def set_format(self, format_type: PageNumberFormat):
        """Set the format of page numbers."""
        if self.settings.format != format_type:
            self.settings.format = format_type
            self._update_page_numbers()
            self.pageNumbersChanged.emit()
    
    def set_start_from(self, start: int):
        """Set the starting page number."""
        if self.settings.start_from != start:
            self.settings.start_from = max(1, start)
            self._update_page_numbers()
            self.pageNumbersChanged.emit()
    
    def set_show_on_first_page(self, show: bool):
        """Set whether to show page numbers on the first page."""
        if self.settings.show_on_first_page != show:
            self.settings.show_on_first_page = show
            self._update_page_numbers()
            self.pageNumbersChanged.emit()
    
    def set_format_string(self, position: PageNumberPosition, format_str: str):
        """Set a custom format string for a specific position."""
        if position in self.settings.format_strings and self.settings.format_strings[position] != format_str:
            self.settings.format_strings[position] = format_str
            self._update_page_numbers()
            self.pageNumbersChanged.emit()
    
    def _on_contents_changed(self):
        """Handle document content changes."""
        self._update_page_numbers()
    
    def _update_page_numbers(self):
        """Update page numbers in the document."""
        if not self.settings.enabled:
            self._remove_page_numbers()
            return
        
        # Calculate total pages
        self._total_pages = self.document.pageCount()
        
        # Get the format string for the current position
        format_str = self.settings.format_strings.get(self.settings.position, "{number}")
        
        # Update page numbers for each page
        for page_num in range(1, self._total_pages + 1):
            # Skip first page if configured
            if page_num == 1 and not self.settings.show_on_first_page:
                continue
                
            # Get the page rect
            page_rect = self.document.documentLayout().pageBoundingRect(page_num - 1)
            
            # Format the page number
            display_num = self._format_page_number(page_num)
            display_text = format_str.format(number=display_num, total=self._total_pages)
            
            # Get the position for the page number
            x, y = self._get_page_number_position(page_rect, display_text)
            
            # Create or update the page number frame
            self._update_page_number_frame(page_num, display_text, QPointF(x, y))
    
    def _format_page_number(self, page_num: int) -> str:
        """Format a page number according to the current format."""
        num = self.settings.start_from + page_num - 1
        
        if self.settings.format == PageNumberFormat.NUMERIC:
            return str(num)
        elif self.settings.format == PageNumberFormat.LOWER_ROMAN:
            return self._to_roman(num).lower()
        elif self.settings.format == PageNumberFormat.UPPER_ROMAN:
            return self._to_roman(num)
        elif self.settings.format == PageNumberFormat.LOWER_ALPHA:
            return self._to_alpha(num).lower()
        elif self.settings.format == PageNumberFormat.UPPER_ALPHA:
            return self._to_alpha(num)
        return str(num)
    
    @staticmethod
    def _to_roman(n: int) -> str:
        """Convert an integer to a Roman numeral."""
        val = [
            1000, 900, 500, 400,
            100, 90, 50, 40,
            10, 9, 5, 4, 1
        ]
        syb = [
            "M", "CM", "D", "CD",
            "C", "XC", "L", "XL",
            "X", "IX", "V", "IV",
            "I"
        ]
        roman_num = ''
        i = 0
        while n > 0:
            for _ in range(n // val[i]):
                roman_num += syb[i]
                n -= val[i]
            i += 1
        return roman_num
    
    @staticmethod
    def _to_alpha(n: int) -> str:
        """Convert an integer to an alphabetic representation (A, B, ..., Z, AA, AB, ...)."""
        result = ""
        n -= 1  # Make it 0-based
        
        while n >= 0:
            remainder = n % 26
            result = chr(65 + remainder) + result
            n = (n // 26) - 1
            
        return result if result else "A"
    
    def _get_page_number_position(self, page_rect: QRectF, text: str) -> Tuple[float, float]:
        """Calculate the position for a page number based on the current settings."""
        # Default margins
        margin = 20  # 20pt margin from edges
        
        # Calculate y position (top or bottom)
        if self.settings.position in [PageNumberPosition.TOP_LEFT, 
                                    PageNumberPosition.TOP_CENTER, 
                                    PageNumberPosition.TOP_RIGHT]:
            y = page_rect.top() + margin
        else:  # BOTTOM_LEFT, BOTTOM_CENTER, BOTTOM_RIGHT, PAGE_X_OF_Y
            y = page_rect.bottom() - margin
        
        # Calculate x position (left, center, or right)
        if self.settings.position in [PageNumberPosition.TOP_LEFT, 
                                    PageNumberPosition.BOTTOM_LEFT]:
            x = page_rect.left() + margin
        elif self.settings.position in [PageNumberPosition.TOP_RIGHT, 
                                      PageNumberPosition.BOTTOM_RIGHT]:
            x = page_rect.right() - margin - self._text_width(text)
        else:  # CENTER or PAGE_X_OF_Y
            x = page_rect.center().x() - (self._text_width(text) / 2)
        
        return x, y
    
    def _text_width(self, text: str) -> float:
        """Calculate the width of text using the document's default font."""
        font = self.document.defaultFont()
        fm = font.metrics()
        return fm.horizontalAdvance(text) * 1.1  # Add 10% for safety
    
    def _update_page_number_frame(self, page_num: int, text: str, position: QPointF):
        """Update or create a frame for the page number."""
        frame_name = f"page_number_{page_num}"
        frame = self._find_or_create_frame(frame_name)
        
        # Set frame format
        frame_format = frame.frameFormat()
        frame_format.setPosition(QTextFrameFormat.Position.FloatLeft)
        frame_format.setBorder(0)
        frame_format.setMargin(0)
        frame_format.setPadding(0)
        frame_format.setPageBreakPolicy(QTextFormat.PageBreak_AlwaysBefore)
        frame_format.setPosition(QTextFrameFormat.Position.FloatLeft)
        frame.setFrameFormat(frame_format)
        
        # Set frame position using absolute positioning
        frame_format.setPosition(QTextFrameFormat.Position.FloatLeft)
        frame_format.setLeftMargin(position.x())
        frame_format.setTopMargin(position.y())
        frame.setFrameFormat(frame_format)
        
        # Update frame content
        cursor = QTextCursor(frame)
        cursor.select(QTextCursor.SelectionType.Document)
        cursor.removeSelectedText()
        
        # Insert page number text
        cursor.insertText(text)
        
        # Apply formatting
        char_format = QTextCharFormat()
        char_format.setFontPointSize(10)  # Default size
        cursor.mergeCharFormat(char_format)
        
        # Align the text
        block_format = QTextBlockFormat()
        if self.settings.position in [PageNumberPosition.TOP_LEFT, 
                                    PageNumberPosition.BOTTOM_LEFT]:
            block_format.setAlignment(Qt.AlignmentFlag.AlignLeft)
        elif self.settings.position in [PageNumberPosition.TOP_RIGHT, 
                                      PageNumberPosition.BOTTOM_RIGHT]:
            block_format.setAlignment(Qt.AlignmentFlag.AlignRight)
        else:  # CENTER or PAGE_X_OF_Y
            block_format.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        cursor.mergeBlockFormat(block_format)
    
    def _find_or_create_frame(self, frame_name: str) -> QTextFrame:
        """Find an existing frame by name or create a new one."""
        root_frame = self.document.rootFrame()
        
        # Try to find existing frame
        child = root_frame.begin()
        while child != root_frame.end():
            frame = child.currentFrame()
            if frame and frame.document().property("frameType") == frame_name:
                return frame
            child += 1
        
        # Create a new frame if not found
        cursor = QTextCursor(self.document)
        frame_format = QTextFrameFormat()
        frame_format.setBorder(0)
        frame_format.setMargin(0)
        frame_format.setPadding(0)
        
        # Create the frame
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        frame = cursor.insertFrame(frame_format)
        frame.document().setProperty("frameType", frame_name)
        
        return frame
    
    def _remove_page_numbers(self):
        """Remove all page number frames from the document."""
        root_frame = self.document.rootFrame()
        cursor = QTextCursor(self.document)
        
        # Find and remove all page number frames
        child = root_frame.begin()
        while child != root_frame.end():
            frame = child.currentFrame()
            if frame and frame.document().property("frameType", "").startswith("page_number_"):
                cursor.setPosition(frame.firstPosition())
                cursor.setPosition(frame.lastPosition(), QTextCursor.MoveMode.KeepAnchor)
                cursor.removeSelectedText()
                child = root_frame.begin()  # Reset iterator after modification
            else:
                child += 1
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert page number settings to a dictionary for serialization."""
        return self.settings.to_dict()
    
    def from_dict(self, data: Dict[str, Any]):
        """Load page number settings from a dictionary."""
        self.settings = PageNumberSettings.from_dict(data)
        self._update_page_numbers()
        self.pageNumbersChanged.emit()

class PageNumberDialog(QDialog):
    """Dialog for configuring page numbers."""
    def __init__(self, page_number_manager: PageNumberManager, parent=None):
        super().__init__(parent)
        self.page_number_manager = page_number_manager
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("Page Numbers")
        self.setMinimumSize(400, 300)
        
        layout = QVBoxLayout()
        
        # Enable/disable page numbers
        self.enable_checkbox = QCheckBox("Show page numbers")
        self.enable_checkbox.setChecked(self.page_number_manager.settings.enabled)
        self.enable_checkbox.toggled.connect(self._on_enable_toggled)
        layout.addWidget(self.enable_checkbox)
        
        # Position
        position_layout = QHBoxLayout()
        position_layout.addWidget(QLabel("Position:"))
        
        self.position_combo = QComboBox()
        for position in PageNumberPosition:
            self.position_combo.addItem(
                position.name.replace('_', ' ').title(), 
                position
            )
        
        # Set current position
        current_index = self.position_combo.findData(self.page_number_manager.settings.position)
        if current_index >= 0:
            self.position_combo.setCurrentIndex(current_index)
        
        self.position_combo.currentIndexChanged.connect(self._on_position_changed)
        position_layout.addWidget(self.position_combo)
        layout.addLayout(position_layout)
        
        # Format
        format_layout = QHBoxLayout()
        format_layout.addWidget(QLabel("Number format:"))
        
        self.format_combo = QComboBox()
        for fmt in PageNumberFormat:
            self.format_combo.addItem(fmt.value, fmt)
        
        # Set current format
        fmt_index = self.format_combo.findData(self.page_number_manager.settings.format)
        if fmt_index >= 0:
            self.format_combo.setCurrentIndex(fmt_index)
        
        self.format_combo.currentIndexChanged.connect(self._on_format_changed)
        format_layout.addWidget(self.format_combo)
        layout.addLayout(format_layout)
        
        # Start from
        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("Start page numbers at:"))
        
        self.start_spin = QSpinBox()
        self.start_spin.setRange(1, 9999)
        self.start_spin.setValue(self.page_number_manager.settings.start_from)
        self.start_spin.valueChanged.connect(self._on_start_changed)
        start_layout.addWidget(self.start_spin)
        layout.addLayout(start_layout)
        
        # Show on first page
        self.first_page_checkbox = QCheckBox("Show on first page")
        self.first_page_checkbox.setChecked(self.page_number_manager.settings.show_on_first_page)
        self.first_page_checkbox.toggled.connect(self._on_first_page_toggled)
        layout.addWidget(self.first_page_checkbox)
        
        # Preview
        preview_label = QLabel("<b>Preview:</b> " + self._get_preview_text())
        layout.addWidget(preview_label)
        
        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
        
        # Connect signals for preview updates
        self.position_combo.currentIndexChanged.connect(
            lambda: preview_label.setText("<b>Preview:</b> " + self._get_preview_text()))
        self.format_combo.currentIndexChanged.connect(
            lambda: preview_label.setText("<b>Preview:</b> " + self._get_preview_text()))
    
    def _get_preview_text(self) -> str:
        """Get a preview of the page number format."""
        position = self.position_combo.currentData()
        fmt = self.format_combo.currentData()
        
        if position == PageNumberPosition.PAGE_X_OF_Y:
            return "Page 1 of 1"
        
        if fmt == PageNumberFormat.NUMERIC:
            return "1"
        elif fmt == PageNumberFormat.LOWER_ROMAN:
            return "i"
        elif fmt == PageNumberFormat.UPPER_ROMAN:
            return "I"
        elif fmt == PageNumberFormat.LOWER_ALPHA:
            return "a"
        elif fmt == PageNumberFormat.UPPER_ALPHA:
            return "A"
        
        return "1"
    
    def _on_enable_toggled(self, enabled: bool):
        """Handle enable/disable of page numbers."""
        self.page_number_manager.set_enabled(enabled)
        
        # Enable/disable other controls
        for widget in [self.position_combo, self.format_combo, 
                      self.start_spin, self.first_page_checkbox]:
            widget.setEnabled(enabled)
    
    def _on_position_changed(self, index: int):
        """Handle position change."""
        position = self.position_combo.itemData(index)
        if position:
            self.page_number_manager.set_position(position)
    
    def _on_format_changed(self, index: int):
        """Handle format change."""
        fmt = self.format_combo.itemData(index)
        if fmt:
            self.page_number_manager.set_format(fmt)
    
    def _on_start_changed(self, value: int):
        """Handle start page number change."""
        self.page_number_manager.set_start_from(value)
    
    def _on_first_page_toggled(self, checked: bool):
        """Handle show on first page toggle."""
        self.page_number_manager.set_show_on_first_page(checked)
