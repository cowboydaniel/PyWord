from PySide6.QtWidgets import (QTextEdit, QApplication, QMenu, QInputDialog,
                             QMessageBox, QFileDialog, QColorDialog)
from PySide6.QtGui import (QTextCharFormat, QTextCursor, QTextBlockFormat,
                         QTextListFormat, QFont, QTextFormat, QTextTableFormat,
                         QTextTable, QTextImageFormat, QTextFrameFormat, QColor,
                         QAction, QTextDocument, QPixmap, QImage, QTextDocumentFragment)
from PySide6.QtCore import Qt, Signal, QSize, QUrl, QRegularExpression
import os
import mimetypes

# Import features
from ..features.styles import DocumentStyles
from ..features.tables import TableManager

class TextEditor(QTextEdit):
    """Advanced text editor with rich text editing capabilities."""
    
    # Signals
    cursorPositionChanged = Signal()
    textChanged = Signal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.zoom_factor = 1.0
        self.current_file = None
        self.styles = DocumentStyles()
        self.table_manager = TableManager(self)
        self.setup_editor()
        self.setup_connections()
        self.setup_context_menu()
    
    def setup_editor(self):
        """Initialize editor settings."""
        self.setAcceptRichText(True)
        self.setLineWrapMode(QTextEdit.WidgetWidth)
        self.setTabStopDistance(40)  # 8 spaces at 12pt font

        # Set default font and style (changed to Calibri for Word compatibility)
        font = QFont("Calibri", 11)
        self.setFont(font)
        
        # Enable undo/redo
        self.setUndoRedoEnabled(True)
        
        # Enable document title changes
        self.document().setModified(False)
        self.document().modificationChanged.connect(self.modification_changed)
        
        # Set up page settings
        self.setup_page_format()
    
    def setup_connections(self):
        """Connect signals to slots."""
        self.cursorPositionChanged.connect(self.update_format)
        self.textChanged.connect(self.on_text_changed)
        
        # Connect table manager signals
        self.cursorPositionChanged.connect(self.update_table_actions)
        
    def setup_context_menu(self):
        """Set up the context menu with formatting options."""
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        
    def show_context_menu(self, position):
        """Show the context menu at the given position."""
        menu = QMenu()
        
        # Standard edit actions
        undo_action = menu.addAction("Undo")
        undo_action.triggered.connect(self.undo)
        undo_action.setEnabled(self.document().isUndoAvailable())
        
        redo_action = menu.addAction("Redo")
        redo_action.triggered.connect(self.redo)
        redo_action.setEnabled(self.document().isRedoAvailable())
        
        menu.addSeparator()
        
        # Formatting actions
        bold_action = menu.addAction("Bold")
        bold_action.setCheckable(True)
        bold_action.setChecked(self.fontWeight() > QFont.Weight.Normal)
        bold_action.triggered.connect(self.text_bold)
        
        italic_action = menu.addAction("Italic")
        italic_action.setCheckable(True)
        italic_action.setChecked(self.fontItalic())
        italic_action.triggered.connect(self.text_italic)
        
        underline_action = menu.addAction("Underline")
        underline_action.setCheckable(True)
        underline_action.setChecked(self.fontUnderline())
        underline_action.triggered.connect(self.text_underline)
        
        # Add more formatting options as needed...
        
        # Show the menu
        menu.exec(self.viewport().mapToGlobal(position))
    
    def update_table_actions(self):
        """Update table-related actions based on cursor position."""
        cursor = self.textCursor()
        in_table = cursor.currentTable() is not None
        
        # Enable/disable table actions based on cursor position
        if hasattr(self, 'insert_row_action'):
            self.insert_row_action.setEnabled(in_table)
            self.delete_row_action.setEnabled(in_table)
            self.insert_column_action.setEnabled(in_table)
            self.delete_column_action.setEnabled(in_table)
    
    def setup_page_format(self):
        """Set up the default page format."""
        # Set up page format (margins, size, etc.)
        text_document = self.document()

        # Set page size to A4 (210 x 297 mm)
        page_size = text_document.pageSize()
        page_size.setWidth(595)  # 210mm at 72dpi
        page_size.setHeight(842)  # 297mm at 72dpi
        text_document.setPageSize(page_size)

        # Set margins (1 inch = 72 points)
        frame_format = QTextFrameFormat()
        frame_format.setLeftMargin(72)  # 1 inch
        frame_format.setRightMargin(72)
        frame_format.setTopMargin(72)
        frame_format.setBottomMargin(72)

        # Apply to the root frame
        root_frame = text_document.rootFrame()
        root_frame.setFrameFormat(frame_format)
    
    def modification_changed(self, changed):
        """Handle document modification state changes."""
        if hasattr(self, 'document_modified'):
            self.document_modified.emit(changed)
    
    def set_document_modified(self, modified):
        """Set the document's modified state."""
        self.document().setModified(modified)
    
    def is_document_modified(self):
        """Return whether the document has been modified."""
        return self.document().isModified()
    
    def get_document_title(self):
        """Get the document title (filename or 'Untitled')."""
        if self.current_file:
            return os.path.basename(self.current_file)
        return "Untitled"
    
    def on_text_changed(self):
        """Handle text changes."""
        # Update word and character count
        text = self.toPlainText()
        word_count = len(text.split())
        char_count = len(text)
        
        # Update status bar if available
        if hasattr(self, 'update_status_bar'):
            self.update_status_bar(word_count, char_count)
        
        # Update document map if available
        if hasattr(self, 'document_map'):
            self.document_map.update_map(0, 0, 0)
        
    def zoom_in(self):
        """Zoom in the document view."""
        self.zoom_factor = min(4.0, self.zoom_factor + 0.1)
        self.update_zoom()
    
    def zoom_out(self):
        """Zoom out the document view."""
        self.zoom_factor = max(0.25, self.zoom_factor - 0.1)
        self.update_zoom()
    
    def zoom_reset(self):
        """Reset zoom to 100%."""
        self.zoom_factor = 1.0
        self.update_zoom()
    
    def set_zoom(self, factor: float):
        """Set zoom to a specific factor.
        
        Args:
            factor: Zoom factor (0.25 to 4.0)
        """
        self.zoom_factor = max(0.25, min(4.0, factor))
        self.update_zoom()
    
    def update_zoom(self):
        """Update the zoom level of the editor."""
        # Scale the font size
        font = self.font()
        base_size = 11  # Default font size at 100% (Calibri 11pt)
        font.setPointSizeF(base_size * self.zoom_factor)
        self.setFont(font)
        
        # Update scroll bars
        self.updateGeometry()
        
        # Notify about zoom change
        if hasattr(self.parent(), 'zoom_changed'):
            self.parent().zoom_changed.emit(int(self.zoom_factor * 100))
    
    # Text formatting methods
    def set_font_family(self, font_family: str):
        """Set font family for selected text or cursor position."""
        fmt = QTextCharFormat()
        fmt.setFontFamily(font_family)
        self.merge_format_on_word_or_selection(fmt)
    
    def set_font_size(self, point_size: int):
        """Set font size for selected text or cursor position."""
        fmt = QTextCharFormat()
        fmt.setFontPointSize(point_size)
        self.merge_format_on_word_or_selection(fmt)
    
    def set_bold(self, enabled: bool):
        """Toggle bold formatting."""
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Bold if enabled else QFont.Normal)
        self.merge_format_on_word_or_selection(fmt)
    
    def set_italic(self, enabled: bool):
        """Toggle italic formatting."""
        fmt = QTextCharFormat()
        fmt.setFontItalic(enabled)
        self.merge_format_on_word_or_selection(fmt)
    
    def set_underline(self, enabled: bool):
        """Toggle underline formatting."""
        fmt = QTextCharFormat()
        fmt.setFontUnderline(enabled)
        self.merge_format_on_word_or_selection(fmt)

    def text_bold(self):
        """Toggle bold formatting on selected text."""
        current_weight = self.fontWeight()
        new_weight = QFont.Weight.Normal if current_weight > QFont.Weight.Normal else QFont.Weight.Bold
        fmt = QTextCharFormat()
        fmt.setFontWeight(new_weight)
        self.merge_format_on_word_or_selection(fmt)

    def text_italic(self):
        """Toggle italic formatting on selected text."""
        fmt = QTextCharFormat()
        fmt.setFontItalic(not self.fontItalic())
        self.merge_format_on_word_or_selection(fmt)

    def text_underline(self):
        """Toggle underline formatting on selected text."""
        fmt = QTextCharFormat()
        fmt.setFontUnderline(not self.fontUnderline())
        self.merge_format_on_word_or_selection(fmt)

    def set_text_color(self, color: QColor):
        """Set text color."""
        fmt = QTextCharFormat()
        fmt.setForeground(color)
        self.merge_format_on_word_or_selection(fmt)
    
    def set_highlight_color(self, color: QColor):
        """Set text background color (highlight)."""
        fmt = QTextCharFormat()
        fmt.setBackground(color)
        self.merge_format_on_word_or_selection(fmt)
    
    # Alignment and formatting
    def set_alignment(self, alignment: Qt.AlignmentFlag):
        """Set paragraph alignment."""
        cursor = self.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.LineUnderCursor)
        
        block_fmt = cursor.blockFormat()
        block_fmt.setAlignment(alignment)
        cursor.mergeBlockFormat(block_fmt)
        self.setTextCursor(cursor)
    
    def set_line_spacing(self, spacing: float):
        """Set line spacing."""
        cursor = self.textCursor()
        block_fmt = cursor.blockFormat()
        block_fmt.setLineHeight(spacing * 100, QTextBlockFormat.ProportionalHeight)
        cursor.mergeBlockFormat(block_fmt)
    
    # Lists
    def insert_bullet_list(self):
        """Insert or toggle bullet list."""
        cursor = self.textCursor()
        cursor.beginEditBlock()
        
        # Create list format if it doesn't exist
        fmt = QTextListFormat()
        if cursor.currentList():
            fmt = cursor.currentList().format()
            cursor.createList(fmt)
        else:
            fmt.setStyle(QTextListFormat.ListDisc)
            cursor.createList(fmt)
        
        cursor.endEditBlock()
    
    def insert_numbered_list(self):
        """Insert or toggle numbered list."""
        cursor = self.textCursor()
        cursor.beginEditBlock()
        
        # Create list format if it doesn't exist
        fmt = QTextListFormat()
        if cursor.currentList():
            fmt = cursor.currentList().format()
            cursor.createList(fmt)
        else:
            fmt.setStyle(QTextListFormat.ListDecimal)
            cursor.createList(fmt)
        
        cursor.endEditBlock()
    
    # Tables
    def insert_table(self, rows: int, columns: int):
        """Insert a table at the current cursor position."""
        cursor = self.textCursor()
        table_format = QTextTableFormat()
        table_format.setCellSpacing(0)
        table_format.setCellPadding(2)
        table_format.setBorder(1)
        
        cursor.insertTable(rows, columns, table_format)
    
    def resize_table(self, rows: int, columns: int):
        """Resize the current table."""
        cursor = self.textCursor()
        table = cursor.currentTable()
        if table:
            table.resize(rows, columns)
    
    # Images
    def insert_image(self, image_path: str, width: int = 200, height: int = 200):
        """Insert an image at the current cursor position."""
        if not os.path.exists(image_path):
            return False
        
        cursor = self.textCursor()
        image_format = QTextImageFormat()
        image_format.setWidth(width)
        image_format.setHeight(height)
        image_format.setName(image_path)
        
        cursor.insertImage(image_format)
        return True
    
    # Helper methods
    def merge_format_on_word_or_selection(self, format: QTextCharFormat):
        """Apply formatting to selected text or word under cursor."""
        cursor = self.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.WordUnderCursor)
        
        cursor.mergeCharFormat(format)
        self.mergeCurrentCharFormat(format)
    
    def update_format(self):
        """Update UI to reflect current text format."""
        # This would be connected to update the formatting toolbar
        # based on the current cursor position
        pass
    
    # Document operations
    def word_count(self) -> dict:
        """Get word, character, and paragraph counts."""
        text = self.toPlainText()
        words = len(text.split())
        chars = len(text)
        paragraphs = len(text.split('\n'))
        
        return {
            'words': words,
            'characters': chars,
            'paragraphs': paragraphs,
            'characters_no_spaces': len(text.replace(' ', ''))
        }
    
    def find_text(self, text: str, options: dict = None) -> bool:
        """Find text in the document."""
        options = options or {}
        flags = QTextDocument.FindFlag(0)
        
        if options.get('case_sensitive', False):
            flags |= QTextDocument.FindCaseSensitively
        if options.get('whole_words', False):
            flags |= QTextDocument.FindWholeWords
        
        return self.find(text, flags)
    
    def replace_text(self, find_text: str, replace_text: str, options: dict = None):
        """Replace text in the document."""
        options = options or {}
        cursor = self.textCursor()
        cursor.beginEditBlock()
        
        # If there's a selection, replace it
        if cursor.hasSelection() and cursor.selectedText() == find_text:
            cursor.insertText(replace_text)
        
        # Find and replace all occurrences
        self.moveCursor(QTextCursor.Start)
        while self.find_text(find_text, options):
            cursor = self.textCursor()
            cursor.insertText(replace_text)
        
        cursor.endEditBlock()
    
    # Undo/redo
    def undo(self):
        """Undo the last operation."""
        self.document().undo()
    
    def redo(self):
        """Redo the last undone operation."""
        self.document().redo()
    
    # Clipboard operations
    def cut(self):
        """Cut selected text to clipboard."""
        super().cut()

    def copy(self):
        """Copy selected text to clipboard."""
        super().copy()

    def paste(self):
        """Paste text from clipboard."""
        super().paste()
    
    def select_all(self):
        """Select all text in the document."""
        self.selectAll()
