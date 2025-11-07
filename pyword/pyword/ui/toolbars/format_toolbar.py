from PySide6.QtWidgets import QToolBar, QComboBox, QFontComboBox, QSpinBox, QToolButton
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QTextCharFormat, QFont, QAction, QIcon

class FormatToolBar(QToolBar):
    def __init__(self, parent=None):
        super().__init__("Format Toolbar", parent)
        self.setIconSize(QSize(16, 16))
        self.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.setup_ui()
    
    def setup_ui(self):
        # Font family
        self.font_family = QFontComboBox(self)
        self.font_family.setEditable(True)
        self.font_family.setCurrentFont(QFont("Arial"))
        self.font_family.setToolTip("Font")
        self.font_family.setFixedWidth(150)
        self.addWidget(self.font_family)
        
        # Font size
        self.font_size = QComboBox(self)
        self.font_size.setEditable(True)
        self.font_size.addItems(["8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"])
        self.font_size.setCurrentText("12")
        self.font_size.setToolTip("Font Size")
        self.font_size.setFixedWidth(60)
        self.addWidget(self.font_size)
        self.addSeparator()
        
        # Bold
        self.bold_action = QAction("Bold", self)
        self.bold_action.setIcon(self.style().standardIcon("SP_DialogApplyButton"))
        self.bold_action.setCheckable(True)
        self.bold_action.setShortcut("Ctrl+B")
        self.bold_action.setToolTip("Bold (Ctrl+B)")
        
        # Italic
        self.italic_action = QAction("Italic", self)
        self.italic_action.setIcon(self.style().standardIcon("SP_DialogResetButton"))
        self.italic_action.setCheckable(True)
        self.italic_action.setShortcut("Ctrl+I")
        self.italic_action.setToolTip("Italic (Ctrl+I)")
        
        # Underline
        self.underline_action = QAction("Underline", self)
        self.underline_action.setIcon(self.style().standardIcon("SP_DialogHelpButton"))
        self.underline_action.setCheckable(True)
        self.underline_action.setShortcut("Ctrl+U")
        self.underline_action.setToolTip("Underline (Ctrl+U)")
        
        # Strikethrough
        self.strikethrough_action = QAction("Strikethrough", self)
        self.strikethrough_action.setIcon(self.style().standardIcon("SP_DialogDiscardButton"))
        self.strikethrough_action.setCheckable(True)
        self.strikethrough_action.setToolTip("Strikethrough")
        
        # Add formatting actions
        self.addAction(self.bold_action)
        self.addAction(self.italic_action)
        self.addAction(self.underline_action)
        self.addAction(self.strikethrough_action)
        self.addSeparator()
        
        # Text color
        self.text_color_action = QAction("Text Color", self)
        self.text_color_action.setIcon(self.create_color_icon(Qt.black))
        self.text_color_action.setToolTip("Text Color")
        
        # Highlight color
        self.highlight_action = QAction("Highlight", self)
        self.highlight_action.setIcon(self.create_highlight_icon(Qt.yellow))
        self.highlight_action.setToolTip("Text Highlight Color")
        
        # Add color actions
        self.addAction(self.text_color_action)
        self.addAction(self.highlight_action)
        self.addSeparator()
        
        # Align Left
        self.align_left_action = QAction("Align Left", self)
        self.align_left_action.setIcon(self.style().standardIcon("SP_ArrowLeft"))
        self.align_left_action.setCheckable(True)
        self.align_left_action.setToolTip("Align Left (Ctrl+L)")
        
        # Align Center
        self.align_center_action = QAction("Align Center", self)
        self.align_center_action.setIcon(self.style().standardIcon("SP_ArrowDown"))
        self.align_center_action.setCheckable(True)
        self.align_center_action.setToolTip("Center (Ctrl+E)")
        
        # Align Right
        self.align_right_action = QAction("Align Right", self)
        self.align_right_action.setIcon(self.style().standardIcon("SP_ArrowRight"))
        self.align_right_action.setCheckable(True)
        self.align_right_action.setToolTip("Align Right (Ctrl+R)")
        
        # Justify
        self.align_justify_action = QAction("Justify", self)
        self.align_justify_action.setIcon(self.style().standardIcon("SP_ArrowUp"))
        self.align_justify_action.setCheckable(True)
        self.align_justify_action.setToolTip("Justify (Ctrl+J)")
        
        # Add alignment actions to a button group
        self.align_group = QAction(self)
        self.align_group.setCheckable(False)
        self.align_left_action.setChecked(True)
        
        self.addAction(self.align_left_action)
        self.addAction(self.align_center_action)
        self.addAction(self.align_right_action)
        self.addAction(self.align_justify_action)
        self.addSeparator()
        
        # Bullet list
        self.bullet_list_action = QAction("Bullet List", self)
        self.bullet_list_action.setIcon(self.style().standardIcon("SP_ArrowDown"))
        self.bullet_list_action.setToolTip("Bullet List")
        
        # Numbered list
        self.numbered_list_action = QAction("Numbered List", self)
        self.numbered_list_action.setIcon(self.style().standardIcon("SP_ArrowDown"))
        self.numbered_list_action.setToolTip("Numbered List")
        
        # Add list actions
        self.addAction(self.bullet_list_action)
        self.addAction(self.numbered_list_action)
        self.addSeparator()
        
        # Decrease indent
        self.decrease_indent_action = QAction("Decrease Indent", self)
        self.decrease_indent_action.setIcon(self.style().standardIcon("SP_ArrowLeft"))
        self.decrease_indent_action.setToolTip("Decrease Indent")
        
        # Increase indent
        self.increase_indent_action = QAction("Increase Indent", self)
        self.increase_indent_action.setIcon(self.style().standardIcon("SP_ArrowRight"))
        self.increase_indent_action.setToolTip("Increase Indent")
        
        # Add indent actions
        self.addAction(self.decrease_indent_action)
        self.addAction(self.increase_indent_action)
    
    def create_color_icon(self, color):
        """Create a colored icon for color picker buttons."""
        pixmap = QPixmap(16, 16)
        pixmap.fill(color)
        return QIcon(pixmap)
    
    def create_highlight_icon(self, color):
        """Create a highlight color icon."""
        pixmap = QPixmap(16, 16)
        pixmap.fill(Qt.transparent)
        
        painter = QPainter(pixmap)
        painter.setBrush(color)
        painter.setPen(Qt.black)
        painter.drawRect(0, 0, 15, 15)
        painter.end()
        
        return QIcon(pixmap)
