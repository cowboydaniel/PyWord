from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QSpinBox, QCheckBox, QGroupBox, QColorDialog, QPushButton)
from PySide6.QtGui import QFont, QColor, QTextCharFormat, QFontDatabase
from PySide6.QtCore import Qt, Signal

from .base_dialog import BaseDialog

class FontDialog(BaseDialog):
    """Font dialog for selecting font properties."""
    
    def __init__(self, parent=None, initial_format=None):
        """Initialize the font dialog.
        
        Args:
            parent: Parent widget
            initial_format: Initial QTextCharFormat or dict with font properties
        """
        super().__init__("Font", parent)
        self.font = QFont()
        self.text_color = QColor(Qt.black)
        self.background_color = QColor(Qt.transparent)
        
        # Set initial values if provided
        if initial_format:
            if isinstance(initial_format, QTextCharFormat):
                self.font = initial_format.font()
                self.text_color = initial_format.foreground().color() if initial_format.hasProperty(QTextCharFormat.ForegroundBrush) else QColor(Qt.black)
                self.background_color = initial_format.background().color() if initial_format.hasProperty(QTextCharFormat.BackgroundBrush) else QColor(Qt.transparent)
            elif isinstance(initial_format, dict):
                self.font.setFamily(initial_format.get('font_family', 'Arial'))
                self.font.setPointSize(initial_format.get('font_size', 12))
                self.font.setBold(initial_format.get('bold', False))
                self.font.setItalic(initial_format.get('italic', False))
                self.font.setUnderline(initial_format.get('underline', False))
                self.font.setStrikeOut(initial_format.get('strikeout', False))
                self.text_color = initial_format.get('text_color', QColor(Qt.black))
                self.background_color = initial_format.get('background_color', QColor(Qt.transparent))
    
    def setup_ui(self):
        """Setup the font dialog UI."""
        # Main layout
        main_layout = QVBoxLayout(self.content)
        
        # Font selection
        font_group = QGroupBox("Font")
        font_layout = QVBoxLayout()
        
        # Font family
        font_family_layout = QHBoxLayout()
        font_family_layout.addWidget(QLabel("Font:"))
        self.font_family_combo = QComboBox()
        self.font_family_combo.addItems([f for f in QFontDatabase().families()])
        self.font_family_combo.setCurrentText(self.font.family())
        self.font_family_combo.currentTextChanged.connect(self.update_preview)
        font_family_layout.addWidget(self.font_family_combo, 1)
        font_layout.addLayout(font_family_layout)
        
        # Font style and size
        font_style_layout = QHBoxLayout()
        
        # Font style
        self.font_style_combo = QComboBox()
        self.update_font_styles()
        self.font_style_combo.currentTextChanged.connect(self.update_font_style)
        font_style_layout.addWidget(QLabel("Font style:"))
        font_style_layout.addWidget(self.font_style_combo, 1)
        
        # Font size
        self.font_size_combo = QComboBox()
        self.font_size_combo.setEditable(True)
        self.font_size_combo.addItems(["8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"])
        self.font_size_combo.setCurrentText(str(self.font.pointSize()))
        self.font_size_combo.currentTextChanged.connect(self.update_font_size)
        font_style_layout.addWidget(QLabel("Size:"))
        font_style_layout.addWidget(self.font_size_combo)
        
        font_layout.addLayout(font_style_layout)
        font_group.setLayout(font_layout)
        
        # Effects
        effects_group = QGroupBox("Effects")
        effects_layout = QVBoxLayout()
        
        # Text color
        color_layout = QHBoxLayout()
        color_layout.addWidget(QLabel("Text color:"))
        self.text_color_btn = QPushButton()
        self.text_color_btn.setFixedSize(24, 24)
        self.text_color_btn.setStyleSheet(f"background-color: {self.text_color.name()}; border: 1px solid black;")
        self.text_color_btn.clicked.connect(self.choose_text_color)
        color_layout.addWidget(self.text_color_btn)
        color_layout.addStretch()
        
        # Background color
        color_layout.addWidget(QLabel("Background color:"))
        self.bg_color_btn = QPushButton()
        self.bg_color_btn.setFixedSize(24, 24)
        self.bg_color_btn.setStyleSheet(f"background-color: {self.background_color.name()}; border: 1px solid black;")
        self.bg_color_btn.clicked.connect(self.choose_bg_color)
        color_layout.addWidget(self.bg_color_btn)
        
        effects_layout.addLayout(color_layout)
        
        # Font effects
        effects_check_layout = QHBoxLayout()
        
        self.bold_check = QCheckBox("Bold")
        self.bold_check.setChecked(self.font.bold())
        self.bold_check.toggled.connect(self.update_bold)
        effects_check_layout.addWidget(self.bold_check)
        
        self.italic_check = QCheckBox("Italic")
        self.italic_check.setChecked(self.font.italic())
        self.italic_check.toggled.connect(self.update_italic)
        effects_check_layout.addWidget(self.italic_check)
        
        self.underline_check = QCheckBox("Underline")
        self.underline_check.setChecked(self.font.underline())
        self.underline_check.toggled.connect(self.update_underline)
        effects_check_layout.addWidget(self.underline_check)
        
        self.strikeout_check = QCheckBox("Strikeout")
        self.strikeout_check.setChecked(self.font.strikeOut())
        self.strikeout_check.toggled.connect(self.update_strikeout)
        effects_check_layout.addWidget(self.strikeout_check)
        
        effects_layout.addLayout(effects_check_layout)
        effects_group.setLayout(effects_layout)
        
        # Preview
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()
        self.preview_label = QLabel("AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumHeight(80)
        self.preview_label.setFrameStyle(1)  # Box
        self.update_preview()
        preview_layout.addWidget(self.preview_label)
        preview_group.setLayout(preview_layout)
        
        # Add groups to main layout
        main_layout.addWidget(font_group)
        main_layout.addWidget(effects_group)
        main_layout.addWidget(preview_group)
        main_layout.addStretch()
    
    def update_font_styles(self):
        """Update the font style combo box based on the selected font family."""
        current_style = self.font_style_combo.currentText()
        self.font_style_combo.clear()
        
        family = self.font_family_combo.currentText()
        styles = QFontDatabase().styles(family)
        
        if not styles:
            styles = ["Regular", "Bold", "Italic", "Bold Italic"]
        
        self.font_style_combo.addItems(styles)
        
        # Try to restore the previous style or select Regular
        index = self.font_style_combo.findText(current_style)
        if index >= 0:
            self.font_style_combo.setCurrentIndex(index)
        else:
            self.font_style_combo.setCurrentText("Regular" if "Regular" in styles else styles[0])
    
    def update_preview(self):
        """Update the preview with the current font settings."""
        # Update font
        self.preview_label.setFont(self.font)
        
        # Update colors
        style_sheet = f"""
        QLabel {{
            color: {color};
            background-color: {bg_color};
            padding: 10px;
        }}
        """.format(
            color=self.text_color.name(),
            bg_color=self.background_color.name() if self.background_color.alpha() > 0 else "transparent"
        )
        self.preview_label.setStyleSheet(style_sheet)
    
    def update_font_style(self, style):
        """Update the font style based on the selected style."""
        family = self.font_family_combo.currentText()
        size = self.font.pointSize()
        
        # Create a font with the selected style
        font = QFontDatabase().font(family, style, size)
        
        # Preserve other font properties
        font.setUnderline(self.font.underline())
        font.setStrikeOut(self.font.strikeOut())
        
        self.font = font
        self.update_preview()
    
    def update_font_size(self, size_text):
        """Update the font size."""
        try:
            size = int(size_text)
            if 1 <= size <= 1638:  # Reasonable font size limits
                self.font.setPointSize(size)
                self.update_preview()
        except ValueError:
            pass
    
    def update_bold(self, checked):
        """Update the bold property of the font."""
        self.font.setBold(checked)
        self.update_preview()
    
    def update_italic(self, checked):
        """Update the italic property of the font."""
        self.font.setItalic(checked)
        self.update_preview()
    
    def update_underline(self, checked):
        """Update the underline property of the font."""
        self.font.setUnderline(checked)
        self.update_preview()
    
    def update_strikeout(self, checked):
        """Update the strikeout property of the font."""
        self.font.setStrikeOut(checked)
        self.update_preview()
    
    def choose_text_color(self):
        """Open a color dialog to choose the text color."""
        color = QColorDialog.getColor(self.text_color, self, "Text Color")
        if color.isValid():
            self.text_color = color
            self.text_color_btn.setStyleSheet(f"background-color: {color.name()}; border: 1px solid black;")
            self.update_preview()
    
    def choose_bg_color(self):
        """Open a color dialog to choose the background color."""
        color = QColorDialog.getColor(self.background_color, self, "Background Color", 
                                     options=QColorDialog.ShowAlphaChannel)
        if color.isValid():
            self.background_color = color
            self.bg_color_btn.setStyleSheet(f"background-color: {color.name()}; border: 1px solid black;")
            self.update_preview()
    
    def get_values(self):
        """Get the current font properties as a dictionary."""
        return {
            'font_family': self.font.family(),
            'font_style': self.font_style_combo.currentText(),
            'font_size': self.font.pointSize(),
            'bold': self.font.bold(),
            'italic': self.font.italic(),
            'underline': self.font.underline(),
            'strikeout': self.font.strikeOut(),
            'text_color': self.text_color,
            'background_color': self.background_color
        }
    
    def set_values(self, values):
        """Set the font properties from a dictionary."""
        if 'font_family' in values:
            self.font_family_combo.setCurrentText(values['font_family'])
        if 'font_style' in values:
            self.font_style_combo.setCurrentText(values['font_style'])
        if 'font_size' in values:
            self.font_size_combo.setCurrentText(str(values['font_size']))
        if 'bold' in values:
            self.bold_check.setChecked(values['bold'])
        if 'italic' in values:
            self.italic_check.setChecked(values['italic'])
        if 'underline' in values:
            self.underline_check.setChecked(values['underline'])
        if 'strikeout' in values:
            self.strikeout_check.setChecked(values['strikeout'])
        if 'text_color' in values:
            self.text_color = values['text_color']
            self.text_color_btn.setStyleSheet(f"background-color: {self.text_color.name()}; border: 1px solid black;")
        if 'background_color' in values:
            self.background_color = values['background_color']
            self.bg_color_btn.setStyleSheet(f"background-color: {self.background_color.name()}; border: 1px solid black;")
        
        self.update_preview()
