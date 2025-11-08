from dataclasses import dataclass
from enum import Enum, auto
from typing import Optional, Dict, Any, List
from PySide6.QtCore import QObject, Signal, QRectF, QSizeF, QPointF, Qt
from PySide6.QtGui import QTextDocument, QTextFrame, QTextFrameFormat, QTextCharFormat, QTextCursor, QTextBlockFormat, QTextFormat
from PySide6.QtWidgets import QTextEdit, QWidget, QVBoxLayout, QLineEdit, QComboBox, QLabel, QDialog, QDialogButtonBox, QFormLayout, QCheckBox, QDoubleSpinBox

class HeaderFooterType(Enum):
    HEADER = auto()
    FOOTER = auto()
    FIRST_PAGE_HEADER = auto()
    FIRST_PAGE_FOOTER = auto()
    EVEN_PAGE_HEADER = auto()
    EVEN_PAGE_FOOTER = auto()

@dataclass
class HeaderFooter:
    content: str = ""
    is_linked_to_previous: bool = True
    position: float = 0.0  # Position from top (for header) or bottom (for footer) in mm
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'content': self.content,
            'is_linked_to_previous': self.is_linked_to_previous,
            'position': self.position
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'HeaderFooter':
        return cls(
            content=data.get('content', ''),
            is_linked_to_previous=data.get('is_linked_to_previous', True),
            position=data.get('position', 0.0)
        )

class HeaderFooterManager(QObject):
    headerFooterChanged = Signal()
    
    def __init__(self, document: QTextDocument, parent=None):
        super().__init__(parent)
        self.document = document
        self.headers: Dict[HeaderFooterType, HeaderFooter] = {
            # First page headers/footers should not be linked to previous (they are the first)
            HeaderFooterType.FIRST_PAGE_HEADER: HeaderFooter("", False, 12.7),
            HeaderFooterType.FIRST_PAGE_FOOTER: HeaderFooter("", False, 12.7),
            # Regular and even page headers/footers should be linked to previous by default
            HeaderFooterType.HEADER: HeaderFooter("", True, 12.7),  # Default 1.27cm from top
            HeaderFooterType.FOOTER: HeaderFooter("", True, 12.7),  # Default 1.27cm from bottom
            HeaderFooterType.EVEN_PAGE_HEADER: HeaderFooter("", True, 12.7),
            HeaderFooterType.EVEN_PAGE_FOOTER: HeaderFooter("", True, 12.7)
        }
        # Only initialize document if it's not None
        if self.document is not None:
            self._init_document()

    def _init_document(self):
        """Initialize document with header and footer frames."""
        if self.document is None:
            return
        root_frame = self.document.rootFrame()
        root_format = root_frame.frameFormat()
        
        # Set page margins to make space for headers and footers
        root_format.setTopMargin(50)  # 50pt = ~1.76cm
        root_format.setBottomMargin(50)
        root_frame.setFrameFormat(root_format)
        
        # Create header and footer frames
        self._create_header_footer_frames()
    
    def _create_header_footer_frames(self):
        """Create frames for headers and footers if they don't exist."""
        root_frame = self.document.rootFrame()

        # Create header frame
        header_frame = self._find_or_create_frame(root_frame, "header")
        header_format = header_frame.frameFormat()
        # Use InFlow positioning as Qt doesn't support FloatTop/FloatBottom for frames
        # Headers and footers positioning is managed through margins and custom rendering
        header_format.setPosition(QTextFrameFormat.Position.InFlow)
        # PageBreakPolicy uses PageBreakFlags enum
        try:
            header_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
        except AttributeError:
            # Fallback for different Qt versions
            pass
        header_frame.setFrameFormat(header_format)

        # Create footer frame
        footer_frame = self._find_or_create_frame(root_frame, "footer")
        footer_format = footer_frame.frameFormat()
        footer_format.setPosition(QTextFrameFormat.Position.InFlow)
        try:
            footer_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysAfter)
        except AttributeError:
            # Fallback for different Qt versions
            pass
        footer_frame.setFrameFormat(footer_format)
    
    def _find_or_create_frame(self, parent: QTextFrame, frame_name: str) -> QTextFrame:
        """Find an existing frame by name or create a new one."""
        frame = None
        child = parent.begin()
        
        while child != parent.end():
            if child.currentFrame() and child.currentFrame().document().property("frameType") == frame_name:
                frame = child.currentFrame()
                break
            child += 1
        
        if not frame:
            # Create a new frame
            cursor = QTextCursor(self.document)
            frame_format = QTextFrameFormat()
            frame_format.setBorder(0)
            frame_format.setMargin(0)
            frame_format.setPadding(0)
            
            if frame_name == "header":
                frame_format.setHeight(20)  # 20pt height for header
            else:  # footer
                frame_format.setHeight(20)  # 20pt height for footer
            
            frame = cursor.insertFrame(frame_format)
            frame.document().setProperty("frameType", frame_name)
        
        return frame
    
    def set_header(self, content: str, header_type: HeaderFooterType = HeaderFooterType.HEADER):
        """Set the content of a header."""
        if header_type in [HeaderFooterType.HEADER, HeaderFooterType.FIRST_PAGE_HEADER, HeaderFooterType.EVEN_PAGE_HEADER]:
            self.headers[header_type].content = content
            self._update_header_footer(header_type)
            self.headerFooterChanged.emit()
    
    def set_footer(self, content: str, footer_type: HeaderFooterType = HeaderFooterType.FOOTER):
        """Set the content of a footer."""
        if footer_type in [HeaderFooterType.FOOTER, HeaderFooterType.FIRST_PAGE_FOOTER, HeaderFooterType.EVEN_PAGE_FOOTER]:
            self.headers[footer_type].content = content
            self._update_header_footer(footer_type)
            self.headerFooterChanged.emit()

    def set_header_footer_content(self, hf_type: HeaderFooterType, content: str):
        """Set the content of a header or footer based on type."""
        if hf_type in [HeaderFooterType.HEADER, HeaderFooterType.FIRST_PAGE_HEADER, HeaderFooterType.EVEN_PAGE_HEADER]:
            self.set_header(content, hf_type)
        elif hf_type in [HeaderFooterType.FOOTER, HeaderFooterType.FIRST_PAGE_FOOTER, HeaderFooterType.EVEN_PAGE_FOOTER]:
            self.set_footer(content, hf_type)

    def _update_header_footer(self, header_footer_type: HeaderFooterType):
        """Update the content of a header or footer frame."""
        # Skip update if document is not set yet
        if self.document is None:
            return

        frame_name = "header" if "header" in header_footer_type.name.lower() else "footer"
        root_frame = self.document.rootFrame()
        frame = self._find_or_create_frame(root_frame, frame_name)
        
        cursor = QTextCursor(frame)
        cursor.select(QTextCursor.SelectionType.Document)
        cursor.removeSelectedText()
        
        # Insert header/footer content
        header_footer = self.headers[header_footer_type]
        cursor.insertText(header_footer.content)
        
        # Apply formatting
        block_format = QTextBlockFormat()
        if frame_name == "header":
            block_format.setAlignment(Qt.AlignmentFlag.AlignCenter)
        else:  # footer
            block_format.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        cursor.select(QTextCursor.SelectionType.Document)
        cursor.mergeBlockFormat(block_format)
    
    def set_position(self, position: float, header_footer_type: HeaderFooterType):
        """Set the position of a header or footer from the edge of the page."""
        if header_footer_type in self.headers:
            self.headers[header_footer_type].position = position
            self._update_header_footer_position(header_footer_type)
            self.headerFooterChanged.emit()
    
    def _update_header_footer_position(self, header_footer_type: HeaderFooterType):
        """Update the position of a header or footer frame."""
        # Skip update if document is not set yet
        if self.document is None:
            return

        frame_name = "header" if "header" in header_footer_type.name.lower() else "footer"
        frame = self._find_or_create_frame(self.document.rootFrame(), frame_name)
        
        frame_format = frame.frameFormat()
        if frame_name == "header":
            frame_format.setTopMargin(self.headers[header_footer_type].position * 2.83465)  # mm to pt
        else:  # footer
            frame_format.setBottomMargin(self.headers[header_footer_type].position * 2.83465)  # mm to pt
        
        frame.setFrameFormat(frame_format)
    
    def set_linked_to_previous(self, linked: bool, header_footer_type: HeaderFooterType):
        """Set whether this header/footer is linked to the previous one."""
        if header_footer_type in self.headers:
            self.headers[header_footer_type].is_linked_to_previous = linked
            self.headerFooterChanged.emit()
    
    def get_header_footer(self, header_footer_type: HeaderFooterType) -> Optional[HeaderFooter]:
        """Get the header or footer of the specified type."""
        return self.headers.get(header_footer_type)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert headers and footers to a dictionary for serialization."""
        return {
            'headers': {k.name: v.to_dict() for k, v in self.headers.items()}
        }
    
    def from_dict(self, data: Dict[str, Any]):
        """Load headers and footers from a dictionary."""
        headers_data = data.get('headers', {})
        for k, v in headers_data.items():
            try:
                header_type = HeaderFooterType[k]
                self.headers[header_type] = HeaderFooter.from_dict(v)
            except KeyError:
                continue
        
        # Update the document with the loaded headers/footers
        for header_type in self.headers:
            self._update_header_footer(header_type)
            self._update_header_footer_position(header_type)

class HeaderFooterDialog(QDialog):
    def __init__(self, header_footer_manager: HeaderFooterManager, 
                 header_footer_type: HeaderFooterType, parent=None):
        super().__init__(parent)
        self.header_footer_manager = header_footer_manager
        self.header_footer_type = header_footer_type
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("Edit Header/Footer")
        self.setMinimumSize(500, 300)
        
        layout = QVBoxLayout()
        
        # Type label
        type_label = QLabel(self.header_footer_type.name.replace('_', ' ').title())
        layout.addWidget(type_label)
        
        # Content editor
        self.content_edit = QTextEdit()
        header_footer = self.header_footer_manager.get_header_footer(self.header_footer_type)
        if header_footer:
            self.content_edit.setPlainText(header_footer.content)
        layout.addWidget(self.content_edit)
        
        # Options
        options_layout = QFormLayout()
        
        # Position
        self.position_spin = QDoubleSpinBox()
        self.position_spin.setRange(0, 50)  # 0-50mm
        self.position_spin.setSuffix(" mm")
        self.position_spin.setValue(header_footer.position if header_footer else 12.7)
        options_layout.addRow("Position from edge:", self.position_spin)
        
        # Link to previous
        self.link_checkbox = QCheckBox("Link to previous")
        self.link_checkbox.setChecked(header_footer.is_linked_to_previous if header_footer else False)
        options_layout.addRow(self.link_checkbox)
        
        layout.addLayout(options_layout)
        
        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def accept(self):
        # Save changes
        content = self.content_edit.toPlainText()
        position = self.position_spin.value()
        linked = self.link_checkbox.isChecked()
        
        if "header" in self.header_footer_type.name.lower():
            self.header_footer_manager.set_header(content, self.header_footer_type)
        else:
            self.header_footer_manager.set_footer(content, self.header_footer_type)
        
        self.header_footer_manager.set_position(position, self.header_footer_type)
        self.header_footer_manager.set_linked_to_previous(linked, self.header_footer_type)
        
        super().accept()
