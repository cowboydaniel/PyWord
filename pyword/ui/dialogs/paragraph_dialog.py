from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QSpinBox, QDoubleSpinBox, QGroupBox, QRadioButton, 
                             QButtonGroup, QDialogButtonBox, QCheckBox, QTabWidget)
from PySide6.QtCore import Qt, QMarginsF
from PySide6.QtGui import QTextBlockFormat, QTextLength

from .base_dialog import BaseDialog

class ParagraphDialog(BaseDialog):
    """Paragraph dialog for setting paragraph formatting."""
    
    def __init__(self, parent=None, initial_format=None):
        """Initialize the paragraph dialog.
        
        Args:
            parent: Parent widget
            initial_format: Initial QTextBlockFormat or dict with paragraph properties
        """
        super().__init__("Paragraph", parent)
        
        # Default values
        self.alignment = Qt.AlignLeft | Qt.AlignAbsolute
        self.left_indent = 0.0
        self.right_indent = 0.0
        self.first_line_indent = 0.0
        self.space_before = 0.0
        self.space_after = 0.0
        self.line_spacing = 1.15  # 1.0 = single, 1.5 = 1.5 lines, etc.
        self.line_spacing_type = 1  # 0 = single, 1 = 1.5, 2 = double, 3 = at least, 4 = exactly, 5 = multiple
        self.keep_together = False
        self.keep_with_next = False
        self.page_break_before = False
        self.widow_orphan_control = True
        
        # Set initial values if provided
        if initial_format:
            if isinstance(initial_format, QTextBlockFormat):
                self.alignment = initial_format.alignment()
                self.left_indent = initial_format.leftMargin()
                self.right_indent = initial_format.rightMargin()
                self.first_line_indent = initial_format.textIndent()
                self.space_before = initial_format.topMargin()
                self.space_after = initial_format.bottomMargin()
                self.line_spacing = initial_format.lineHeight() / 100.0 if initial_format.hasProperty(QTextBlockFormat.LineHeight) else 1.15
                self.line_spacing_type = initial_format.lineHeightType() if initial_format.hasProperty(QTextBlockFormat.LineHeightType) else 1
                self.keep_together = initial_format.nonBreakableLines()
                self.keep_with_next = initial_format.pageBreakPolicy() & QTextBlockFormat.PageBreak_KeepWithNext
                self.page_break_before = initial_format.pageBreakPolicy() & QTextBlockFormat.PageBreak_AlwaysBefore
                self.widow_orphan_control = not (initial_format.pageBreakPolicy() & QTextBlockFormat.PageBreak_AlwaysEnabled)
            elif isinstance(initial_format, dict):
                self.alignment = initial_format.get('alignment', Qt.AlignLeft | Qt.AlignAbsolute)
                self.left_indent = initial_format.get('left_indent', 0.0)
                self.right_indent = initial_format.get('right_indent', 0.0)
                self.first_line_indent = initial_format.get('first_line_indent', 0.0)
                self.space_before = initial_format.get('space_before', 0.0)
                self.space_after = initial_format.get('space_after', 0.0)
                self.line_spacing = initial_format.get('line_spacing', 1.15)
                self.line_spacing_type = initial_format.get('line_spacing_type', 1)
                self.keep_together = initial_format.get('keep_together', False)
                self.keep_with_next = initial_format.get('keep_with_next', False)
                self.page_break_before = initial_format.get('page_break_before', False)
                self.widow_orphan_control = initial_format.get('widow_orphan_control', True)
    
    def setup_ui(self):
        """Setup the paragraph dialog UI."""
        # Main layout
        main_layout = QVBoxLayout(self.content)
        
        # Create tabs
        self.tabs = QTabWidget()
        
        # Indents and Spacing tab
        indents_tab = QWidget()
        self.setup_indents_tab(indents_tab)
        self.tabs.addTab(indents_tab, "Indents and Spacing")
        
        # Line and Page Breaks tab
        breaks_tab = QWidget()
        self.setup_breaks_tab(breaks_tab)
        self.tabs.addTab(breaks_tab, "Line and Page Breaks")
        
        main_layout.addWidget(self.tabs)
        
        # Preview
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()
        self.preview_label = QLabel("This is a preview of how the paragraph will look with the selected formatting.")
        self.preview_label.setWordWrap(True)
        self.preview_label.setFrameStyle(1)  # Box
        self.preview_label.setMinimumHeight(80)
        self.update_preview()
        preview_layout.addWidget(self.preview_label)
        preview_group.setLayout(preview_layout)
        
        main_layout.addWidget(preview_group)
        main_layout.addStretch()
    
    def setup_indents_tab(self, parent):
        """Setup the Indents and Spacing tab."""
        layout = QVBoxLayout(parent)
        
        # Alignment
        align_group = QGroupBox("Alignment")
        align_layout = QHBoxLayout()
        
        self.align_group = QButtonGroup(self)
        
        align_btns = [
            ("Left", "text-align: left;", Qt.AlignLeft | Qt.AlignAbsolute),
            ("Center", "text-align: center;", Qt.AlignHCenter),
            ("Right", "text-align: right;", Qt.AlignRight | Qt.AlignAbsolute),
            ("Justified", "text-align: justify;", Qt.AlignJustify)
        ]
        
        for text, style, alignment in align_btns:
            btn = QRadioButton(text)
            btn.setStyleSheet(style)
            btn.clicked.connect(lambda checked, a=alignment: self.update_alignment(a))
            self.align_group.addButton(btn, alignment)
            align_layout.addWidget(btn)
        
        align_group.setLayout(align_layout)
        
        # Set current alignment
        for btn in self.align_group.buttons():
            if self.align_group.id(btn) == self.alignment:
                btn.setChecked(True)
                break
        
        # Indentation
        indent_group = QGroupBox("Indentation")
        indent_layout = QVBoxLayout()
        
        # Left indent
        left_layout = QHBoxLayout()
        left_layout.addWidget(QLabel("Left:"))
        self.left_indent_spin = QDoubleSpinBox()
        self.left_indent_spin.setRange(0, 31680)  # 22 inches in points (22 * 1440)
        self.left_indent_spin.setValue(self.left_indent)
        self.left_indent_spin.setSuffix(" pt")
        self.left_indent_spin.valueChanged.connect(self.update_left_indent)
        left_layout.addWidget(self.left_indent_spin)
        left_layout.addStretch()
        
        # Right indent
        right_layout = QHBoxLayout()
        right_layout.addWidget(QLabel("Right:"))
        self.right_indent_spin = QDoubleSpinBox()
        self.right_indent_spin.setRange(0, 31680)
        self.right_indent_spin.setValue(self.right_indent)
        self.right_indent_spin.setSuffix(" pt")
        self.right_indent_spin.valueChanged.connect(self.update_right_indent)
        right_layout.addWidget(self.right_indent_spin)
        right_layout.addStretch()
        
        # First line indent
        first_line_layout = QHBoxLayout()
        first_line_layout.addWidget(QLabel("First line:"))
        self.first_line_indent_combo = QComboBox()
        self.first_line_indent_combo.addItems(["(none)", "First line", "Hanging"])
        self.first_line_indent_combo.currentIndexChanged.connect(self.update_first_line_indent_type)
        
        self.first_line_indent_spin = QDoubleSpinBox()
        self.first_line_indent_spin.setRange(-31680, 31680)
        self.first_line_indent_spin.setValue(abs(self.first_line_indent))
        self.first_line_indent_spin.setSuffix(" pt")
        self.first_line_indent_spin.valueChanged.connect(self.update_first_line_indent_value)
        
        # Set initial state
        if self.first_line_indent > 0:
            self.first_line_indent_combo.setCurrentIndex(1)  # First line
            self.first_line_indent_spin.setValue(self.first_line_indent)
        elif self.first_line_indent < 0:
            self.first_line_indent_combo.setCurrentIndex(2)  # Hanging
            self.first_line_indent_spin.setValue(-self.first_line_indent)
        
        first_line_layout.addWidget(self.first_line_indent_combo)
        first_line_layout.addWidget(self.first_line_indent_spin)
        first_line_layout.addStretch()
        
        indent_layout.addLayout(left_layout)
        indent_layout.addLayout(right_layout)
        indent_layout.addLayout(first_line_layout)
        indent_group.setLayout(indent_layout)
        
        # Spacing
        spacing_group = QGroupBox("Spacing")
        spacing_layout = QVBoxLayout()
        
        # Before
        before_layout = QHBoxLayout()
        before_layout.addWidget(QLabel("Before:"))
        self.before_spin = QDoubleSpinBox()
        self.before_spin.setRange(0, 1584)  # 11 inches in points (11 * 144)
        self.before_spin.setValue(self.space_before)
        self.before_spin.setSuffix(" pt")
        self.before_spin.valueChanged.connect(self.update_space_before)
        before_layout.addWidget(self.before_spin)
        before_layout.addStretch()
        
        # After
        after_layout = QHBoxLayout()
        after_layout.addWidget(QLabel("After:"))
        self.after_spin = QDoubleSpinBox()
        self.after_spin.setRange(0, 1584)
        self.after_spin.setValue(self.space_after)
        self.after_spin.setSuffix(" pt")
        self.after_spin.valueChanged.connect(self.update_space_after)
        after_layout.addWidget(self.after_spin)
        after_layout.addStretch()
        
        # Line spacing
        line_spacing_layout = QHBoxLayout()
        line_spacing_layout.addWidget(QLabel("Line spacing:"))
        
        self.line_spacing_combo = QComboBox()
        self.line_spacing_combo.addItems(["Single", "1.5 lines", "Double", "At least", "Exactly", "Multiple"])
        self.line_spacing_combo.setCurrentIndex(self.line_spacing_type)
        self.line_spacing_combo.currentIndexChanged.connect(self.update_line_spacing_type)
        
        self.line_spacing_spin = QDoubleSpinBox()
        self.line_spacing_spin.setRange(0.25, 132.0)  # 0.25 to 132 points
        self.line_spacing_spin.setValue(self.line_spacing * 12.0)  # Convert to points for display
        self.line_spacing_spin.setSuffix(" pt")
        self.line_spacing_spin.valueChanged.connect(self.update_line_spacing_value)
        
        # Show/hide the spin box based on line spacing type
        self.update_line_spacing_controls()
        
        line_spacing_layout.addWidget(self.line_spacing_combo)
        line_spacing_layout.addWidget(self.line_spacing_spin)
        line_spacing_layout.addStretch()
        
        spacing_layout.addLayout(before_layout)
        spacing_layout.addLayout(after_layout)
        spacing_layout.addLayout(line_spacing_layout)
        spacing_group.setLayout(spacing_layout)
        
        # Add all to layout
        layout.addWidget(align_group)
        layout.addWidget(indent_group)
        layout.addWidget(spacing_group)
        layout.addStretch()
    
    def setup_breaks_tab(self, parent):
        """Setup the Line and Page Breaks tab."""
        layout = QVBoxLayout(parent)
        
        # Pagination
        pagination_group = QGroupBox("Pagination")
        pagination_layout = QVBoxLayout()
        
        self.widow_orphan_check = QCheckBox("Widow/Orphan control")
        self.widow_orphan_check.setChecked(self.widow_orphan_control)
        self.widow_orphan_check.toggled.connect(self.update_widow_orphan_control)
        
        self.keep_together_check = QCheckBox("Keep lines together")
        self.keep_together_check.setChecked(self.keep_together)
        self.keep_together_check.toggled.connect(self.update_keep_together)
        
        self.keep_with_next_check = QCheckBox("Keep with next")
        self.keep_with_next_check.setChecked(self.keep_with_next)
        self.keep_with_next_check.toggled.connect(self.update_keep_with_next)
        
        self.page_break_before_check = QCheckBox("Page break before")
        self.page_break_before_check.setChecked(self.page_break_before)
        self.page_break_before_check.toggled.connect(self.update_page_break_before)
        
        pagination_layout.addWidget(self.widow_orphan_check)
        pagination_layout.addWidget(self.keep_together_check)
        pagination_layout.addWidget(self.keep_with_next_check)
        pagination_layout.addWidget(self.page_break_before_check)
        pagination_group.setLayout(pagination_layout)
        
        layout.addWidget(pagination_group)
        layout.addStretch()
    
    def update_alignment(self, alignment):
        """Update the paragraph alignment."""
        self.alignment = alignment
        self.update_preview()
    
    def update_left_indent(self, value):
        """Update the left indent."""
        self.left_indent = value
        self.update_preview()
    
    def update_right_indent(self, value):
        """Update the right indent."""
        self.right_indent = value
        self.update_preview()
    
    def update_first_line_indent_type(self, index):
        """Update the first line indent type."""
        if index == 0:  # None
            self.first_line_indent = 0.0
        elif index == 1:  # First line
            self.first_line_indent = self.first_line_indent_spin.value()
        else:  # Hanging
            self.first_line_indent = -self.first_line_indent_spin.value()
        self.update_preview()
    
    def update_first_line_indent_value(self, value):
        """Update the first line indent value."""
        if self.first_line_indent_combo.currentIndex() == 1:  # First line
            self.first_line_indent = value
        elif self.first_line_indent_combo.currentIndex() == 2:  # Hanging
            self.first_line_indent = -value
        self.update_preview()
    
    def update_space_before(self, value):
        """Update the space before the paragraph."""
        self.space_before = value
        self.update_preview()
    
    def update_space_after(self, value):
        """Update the space after the paragraph."""
        self.space_after = value
        self.update_preview()
    
    def update_line_spacing_type(self, index):
        """Update the line spacing type."""
        self.line_spacing_type = index
        self.update_line_spacing_controls()
        self.update_preview()
    
    def update_line_spacing_value(self, value):
        """Update the line spacing value."""
        if self.line_spacing_type in [0, 1, 2]:  # Single, 1.5, Double
            # These are fixed values, so we don't update the line spacing
            pass
        else:
            self.line_spacing = value / 12.0  # Convert from points to lines
            self.update_preview()
    
    def update_line_spacing_controls(self):
        """Update the line spacing controls based on the selected type."""
        show_spin = self.line_spacing_type >= 3  # Show for At least, Exactly, Multiple
        self.line_spacing_spin.setVisible(show_spin)
        
        # Set appropriate value and suffix
        if self.line_spacing_type == 0:  # Single
            self.line_spacing = 1.0
            self.line_spacing_spin.setValue(12.0)
        elif self.line_spacing_type == 1:  # 1.5 lines
            self.line_spacing = 1.5
            self.line_spacing_spin.setValue(18.0)
        elif self.line_spacing_type == 2:  # Double
            self.line_spacing = 2.0
            self.line_spacing_spin.setValue(24.0)
        elif self.line_spacing_type == 3:  # At least
            self.line_spacing_spin.setSuffix(" pt")
            self.line_spacing_spin.setValue(self.line_spacing * 12.0)
        elif self.line_spacing_type == 4:  # Exactly
            self.line_spacing_spin.setSuffix(" pt")
            self.line_spacing_spin.setValue(self.line_spacing * 12.0)
        else:  # Multiple
            self.line_spacing_spin.setSuffix(" lines")
            self.line_spacing_spin.setValue(self.line_spacing)
    
    def update_widow_orphan_control(self, enabled):
        """Update the widow/orphan control setting."""
        self.widow_orphan_control = enabled
        self.update_preview()
    
    def update_keep_together(self, enabled):
        """Update the keep together setting."""
        self.keep_together = enabled
        self.update_preview()
    
    def update_keep_with_next(self, enabled):
        """Update the keep with next setting."""
        self.keep_with_next = enabled
        self.update_preview()
    
    def update_page_break_before(self, enabled):
        """Update the page break before setting."""
        self.page_break_before = enabled
        self.update_preview()
    
    def update_preview(self):
        """Update the preview with the current paragraph formatting."""
        # Create a style string for the preview
        style = []
        
        # Alignment
        if self.alignment & Qt.AlignLeft:
            style.append("text-align: left;")
        elif self.alignment & Qt.AlignHCenter:
            style.append("text-align: center;")
        elif self.alignment & Qt.AlignRight:
            style.append("text-align: right;")
        elif self.alignment & Qt.AlignJustify:
            style.append("text-align: justify;")
        
        # Indentation
        if self.left_indent > 0:
            style.append(f"margin-left: {self.left_indent}pt;")
        if self.right_indent > 0:
            style.append(f"margin-right: {self.right_indent}pt;")
        if self.first_line_indent > 0:
            style.append(f"text-indent: {self.first_line_indent}pt;")
        elif self.first_line_indent < 0:
            style.append(f"text-indent: 0; padding-left: {-self.first_line_indent}pt; text-indent: {-self.first_line_indent}pt;")
        
        # Spacing
        if self.space_before > 0:
            style.append(f"margin-top: {self.space_before}pt;")
        if self.space_after > 0:
            style.append(f"margin-bottom: {self.space_after}pt;")
        
        # Line height
        if self.line_spacing_type == 0:  # Single
            style.append("line-height: 1.0;")
        elif self.line_spacing_type == 1:  # 1.5 lines
            style.append("line-height: 1.5;")
        elif self.line_spacing_type == 2:  # Double
            style.append("line-height: 2.0;")
        elif self.line_spacing_type == 3:  # At least
            style.append(f"line-height: {self.line_spacing}; min-height: {self.line_spacing * 12.0}pt;")
        elif self.line_spacing_type == 4:  # Exactly
            style.append(f"line-height: {self.line_spacing}; height: {self.line_spacing * 12.0}pt;")
        else:  # Multiple
            style.append(f"line-height: {self.line_spacing};")
        
        # Apply the style to the preview
        self.preview_label.setStyleSheet("QLabel { " + " ".join(style) + " }")
    
    def get_values(self):
        """Get the current paragraph properties as a dictionary."""
        return {
            'alignment': self.alignment,
            'left_indent': self.left_indent,
            'right_indent': self.right_indent,
            'first_line_indent': self.first_line_indent,
            'space_before': self.space_before,
            'space_after': self.space_after,
            'line_spacing': self.line_spacing,
            'line_spacing_type': self.line_spacing_type,
            'keep_together': self.keep_together,
            'keep_with_next': self.keep_with_next,
            'page_break_before': self.page_break_before,
            'widow_orphan_control': self.widow_orphan_control
        }
    
    def set_values(self, values):
        """Set the paragraph properties from a dictionary."""
        if 'alignment' in values:
            self.alignment = values['alignment']
            for btn in self.align_group.buttons():
                if self.align_group.id(btn) == self.alignment:
                    btn.setChecked(True)
                    break
        
        if 'left_indent' in values:
            self.left_indent = values['left_indent']
            self.left_indent_spin.setValue(self.left_indent)
        
        if 'right_indent' in values:
            self.right_indent = values['right_indent']
            self.right_indent_spin.setValue(self.right_indent)
        
        if 'first_line_indent' in values:
            self.first_line_indent = values['first_line_indent']
            if self.first_line_indent > 0:
                self.first_line_indent_combo.setCurrentIndex(1)  # First line
                self.first_line_indent_spin.setValue(self.first_line_indent)
            elif self.first_line_indent < 0:
                self.first_line_indent_combo.setCurrentIndex(2)  # Hanging
                self.first_line_indent_spin.setValue(-self.first_line_indent)
            else:
                self.first_line_indent_combo.setCurrentIndex(0)  # None
        
        if 'space_before' in values:
            self.space_before = values['space_before']
            self.before_spin.setValue(self.space_before)
        
        if 'space_after' in values:
            self.space_after = values['space_after']
            self.after_spin.setValue(self.space_after)
        
        if 'line_spacing' in values:
            self.line_spacing = values['line_spacing']
        
        if 'line_spacing_type' in values:
            self.line_spacing_type = values['line_spacing_type']
            self.line_spacing_combo.setCurrentIndex(self.line_spacing_type)
            self.update_line_spacing_controls()
        
        if 'keep_together' in values:
            self.keep_together = values['keep_together']
            self.keep_together_check.setChecked(self.keep_together)
        
        if 'keep_with_next' in values:
            self.keep_with_next = values['keep_with_next']
            self.keep_with_next_check.setChecked(self.keep_with_next)
        
        if 'page_break_before' in values:
            self.page_break_before = values['page_break_before']
            self.page_break_before_check.setChecked(self.page_break_before)
        
        if 'widow_orphan_control' in values:
            self.widow_orphan_control = values['widow_orphan_control']
            self.widow_orphan_check.setChecked(self.widow_orphan_control)
        
        self.update_preview()
