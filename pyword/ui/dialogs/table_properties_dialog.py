from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QSpinBox, QDoubleSpinBox, QGroupBox, QRadioButton, 
                             QButtonGroup, QDialogButtonBox, QCheckBox, QTabWidget,
                             QTableWidget, QTableWidgetItem, QHeaderView, QColorDialog,
                             QPushButton, QLineEdit, QGridLayout, QSizePolicy)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QColor, QBrush, QPen, QPainter, QPixmap

from .base_dialog import BaseDialog

class TablePropertiesDialog(BaseDialog):
    """Table properties dialog for setting table formatting."""
    
    def __init__(self, parent=None, initial_properties=None):
        """Initialize the table properties dialog.
        
        Args:
            parent: Parent widget
            initial_properties: Initial table properties as a dictionary
        """
        super().__init__("Table Properties", parent)
        
        # Default values
        self.properties = {
            'alignment': Qt.AlignLeft | Qt.AlignAbsolute,
            'indent': 0.0,
            'size': {'width': 6.5, 'unit': 'inches', 'percent': 100},
            'text_wrapping': True,
            'borders': {
                'all': {'style': 'single', 'color': QColor(Qt.black), 'width': 0.5},
                'top': {'style': 'single', 'color': QColor(Qt.black), 'width': 0.5},
                'bottom': {'style': 'single', 'color': QColor(Qt.black), 'width': 0.5},
                'left': {'style': 'single', 'color': QColor(Qt.black), 'width': 0.5},
                'right': {'style': 'single', 'color': QColor(Qt.black), 'width': 0.5},
                'inside_h': {'style': 'single', 'color': QColor(Qt.black), 'width': 0.5},
                'inside_v': {'style': 'single', 'color': QColor(Qt.black), 'width': 0.5}
            },
            'shading': {
                'fill': QColor(Qt.transparent),
                'color': QColor(Qt.transparent)
            },
            'cell_margins': {
                'top': 0.0,
                'bottom': 0.0,
                'left': 0.1,
                'right': 0.1
            },
            'cell_spacing': 0.0,
            'column_widths': [],
            'preferred_width': {'type': 'auto', 'value': 100, 'unit': 'percent'},
            'allow_break_across_pages': True,
            'repeat_header_row': False,
            'description': ''
        }
        
        # Update with initial properties if provided
        if initial_properties:
            self.properties.update(initial_properties)
        
        self.setup_ui()
        self.update_preview()
    
    def setup_ui(self):
        """Setup the table properties dialog UI."""
        # Main layout
        main_layout = QHBoxLayout(self.content)
        
        # Left side: Tabs
        self.tabs = QTabWidget()
        
        # Table tab
        table_tab = QWidget()
        self.setup_table_tab(table_tab)
        self.tabs.addTab(table_tab, "Table")
        
        # Row tab
        row_tab = QWidget()
        self.setup_row_tab(row_tab)
        self.tabs.addTab(row_tab, "Row")
        
        # Column tab
        column_tab = QWidget()
        self.setup_column_tab(column_tab)
        self.tabs.addTab(column_tab, "Column")
        
        # Cell tab
        cell_tab = QWidget()
        self.setup_cell_tab(cell_tab)
        self.tabs.addTab(cell_tab, "Cell")
        
        # Borders and Shading tab
        borders_tab = QWidget()
        self.setup_borders_tab(borders_tab)
        self.tabs.addTab(borders_tab, "Borders and Shading")
        
        # Alt Text tab
        alt_text_tab = QWidget()
        self.setup_alt_text_tab(alt_text_tab)
        self.tabs.addTab(alt_text_tab, "Alt Text")
        
        # Add tabs to main layout
        main_layout.addWidget(self.tabs, 1)
        
        # Right side: Preview
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()
        
        self.preview_widget = QWidget()
        self.preview_widget.setMinimumSize(200, 200)
        self.preview_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        preview_layout.addWidget(self.preview_widget)
        preview_group.setLayout(preview_layout)
        
        main_layout.addWidget(preview_group, 1)
    
    def setup_table_tab(self, parent):
        """Setup the Table tab."""
        layout = QVBoxLayout(parent)
        
        # Size group
        size_group = QGroupBox("Size")
        size_layout = QHBoxLayout()
        
        size_layout.addWidget(QLabel("Preferred width:"))
        
        self.width_spin = QDoubleSpinBox()
        self.width_spin.setRange(0.1, 22.0)
        self.width_spin.setValue(self.properties['preferred_width']['value'])
        self.width_spin.setSingleStep(0.1)
        self.width_spin.valueChanged.connect(self.update_table_width)
        
        self.width_unit_combo = QComboBox()
        self.width_unit_combo.addItems(["Percent", "Inches", "Centimeters", "Points"])
        self.width_unit_combo.setCurrentText(self.properties['preferred_width']['unit'].capitalize())
        self.width_unit_combo.currentTextChanged.connect(self.update_table_width_unit)
        
        size_layout.addWidget(self.width_spin)
        size_layout.addWidget(self.width_unit_combo)
        size_group.setLayout(size_layout)
        
        # Alignment group
        align_group = QGroupBox("Alignment")
        align_layout = QVBoxLayout()
        
        self.align_buttons = QButtonGroup()
        
        align_btns = [
            ("Left", "Align Left", Qt.AlignLeft | Qt.AlignAbsolute),
            ("Center", "Center", Qt.AlignHCenter),
            ("Right", "Right", Qt.AlignRight | Qt.AlignAbsolute)
        ]
        
        for text, tooltip, alignment in align_btns:
            btn = QRadioButton(text)
            btn.setToolTip(tooltip)
            btn.clicked.connect(lambda checked, a=alignment: self.update_alignment(a))
            self.align_buttons.addButton(btn, alignment)
            align_layout.addWidget(btn)
        
        # Set current alignment
        for btn in self.align_buttons.buttons():
            if self.align_buttons.id(btn) == self.properties['alignment']:
                btn.setChecked(True)
                break
        
        align_layout.addStretch()
        align_group.setLayout(align_layout)
        
        # Text wrapping
        text_wrap_group = QGroupBox("Text Wrapping")
        text_wrap_layout = QVBoxLayout()
        
        self.wrap_none_radio = QRadioButton("None")
        self.wrap_around_radio = QRadioButton("Around")
        
        if self.properties['text_wrapping']:
            self.wrap_around_radio.setChecked(True)
        else:
            self.wrap_none_radio.setChecked(True)
        
        self.wrap_none_radio.toggled.connect(self.update_text_wrapping)
        
        text_wrap_layout.addWidget(self.wrap_none_radio)
        text_wrap_layout.addWidget(self.wrap_around_radio)
        text_wrap_layout.addStretch()
        text_wrap_group.setLayout(text_wrap_layout)
        
        # Positioning (only enabled when text wrapping is on)
        self.positioning_group = QGroupBox("Positioning")
        positioning_layout = QGridLayout()
        
        positioning_layout.addWidget(QLabel("Horizontal:"), 0, 0)
        self.horizontal_pos_combo = QComboBox()
        self.horizontal_pos_combo.addItems(["Left", "Center", "Right", "Inside", "Outside"])
        positioning_layout.addWidget(self.horizontal_pos_combo, 0, 1)
        
        positioning_layout.addWidget(QLabel("Relative to:"), 0, 2)
        self.horizontal_rel_combo = QComboBox()
        self.horizontal_rel_combo.addItems(["Column", "Margin", "Page"])
        positioning_layout.addWidget(self.horizontal_rel_combo, 0, 3)
        
        positioning_layout.addWidget(QLabel("Vertical:"), 1, 0)
        self.vertical_pos_combo = QComboBox()
        self.vertical_pos_combo.addItems(["Top", "Center", "Bottom", "Inside", "Outside"])
        positioning_layout.addWidget(self.vertical_pos_combo, 1, 1)
        
        positioning_layout.addWidget(QLabel("Relative to:"), 1, 2)
        self.vertical_rel_combo = QComboBox()
        self.vertical_rel_combo.addItems(["Margin", "Page", "Paragraph", "Line"])
        positioning_layout.addWidget(self.vertical_rel_combo, 1, 3)
        
        positioning_layout.addWidget(QLabel("Distance from text:"), 2, 0)
        self.distance_top_spin = QDoubleSpinBox()
        self.distance_top_spin.setRange(0, 31680)  # 22 inches in points
        self.distance_top_spin.setSuffix(" pt")
        positioning_layout.addWidget(QLabel("Top:"), 2, 1)
        positioning_layout.addWidget(self.distance_top_spin, 2, 2)
        
        self.distance_bottom_spin = QDoubleSpinBox()
        self.distance_bottom_spin.setRange(0, 31680)
        self.distance_bottom_spin.setSuffix(" pt")
        positioning_layout.addWidget(QLabel("Bottom:"), 3, 1)
        positioning_layout.addWidget(self.distance_bottom_spin, 3, 2)
        
        self.distance_left_spin = QDoubleSpinBox()
        self.distance_left_spin.setRange(0, 31680)
        self.distance_left_spin.setSuffix(" pt")
        positioning_layout.addWidget(QLabel("Left:"), 4, 1)
        positioning_layout.addWidget(self.distance_left_spin, 4, 2)
        
        self.distance_right_spin = QDoubleSpinBox()
        self.distance_right_spin.setRange(0, 31680)
        self.distance_right_spin.setSuffix(" pt")
        positioning_layout.addWidget(QLabel("Right:"), 5, 1)
        positioning_layout.addWidget(self.distance_right_spin, 5, 2)
        
        self.positioning_group.setLayout(positioning_layout)
        self.positioning_group.setEnabled(self.properties['text_wrapping'])
        
        # Options
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()
        
        self.allow_break_check = QCheckBox("Allow row to break across pages")
        self.allow_break_check.setChecked(self.properties['allow_break_across_pages'])
        self.allow_break_check.toggled.connect(self.update_break_across_pages)
        
        self.repeat_header_check = QCheckBox("Repeat as header row at the top of each page")
        self.repeat_header_check.setChecked(self.properties['repeat_header_row'])
        self.repeat_header_check.toggled.connect(self.update_repeat_header)
        
        options_layout.addWidget(self.allow_break_check)
        options_layout.addWidget(self.repeat_header_check)
        options_layout.addStretch()
        options_group.setLayout(options_layout)
        
        # Add all groups to layout
        layout.addWidget(size_group)
        layout.addWidget(align_group)
        layout.addWidget(text_wrap_group)
        layout.addWidget(self.positioning_group)
        layout.addWidget(options_group)
        layout.addStretch()
    
    def setup_row_tab(self, parent):
        """Setup the Row tab."""
        layout = QVBoxLayout(parent)
        
        # Size group
        size_group = QGroupBox("Size")
        size_layout = QHBoxLayout()
        
        size_layout.addWidget(QLabel("Height:"))
        
        self.row_height_spin = QDoubleSpinBox()
        self.row_height_spin.setRange(0, 31680)  # 22 inches in points
        self.row_height_spin.setValue(0)  # Auto
        self.row_height_spin.setSuffix(" pt")
        
        self.row_height_combo = QComboBox()
        self.row_height_combo.addItems(["At least", "Exactly"])
        
        size_layout.addWidget(self.row_height_spin)
        size_layout.addWidget(self.row_height_combo)
        size_group.setLayout(size_layout)
        
        # Options
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()
        
        self.allow_row_break_check = QCheckBox("Allow row to break across pages")
        self.allow_row_break_check.setChecked(True)
        
        self.repeat_header_row_check = QCheckBox("Repeat as header row at the top of each page")
        
        options_layout.addWidget(self.allow_row_break_check)
        options_layout.addWidget(self.repeat_header_row_check)
        options_layout.addStretch()
        options_group.setLayout(options_layout)
        
        # Add groups to layout
        layout.addWidget(size_group)
        layout.addWidget(options_group)
        layout.addStretch()
    
    def setup_column_tab(self, parent):
        """Setup the Column tab."""
        layout = QVBoxLayout(parent)
        
        # Size group
        size_group = QGroupBox("Size")
        size_layout = QHBoxLayout()
        
        size_layout.addWidget(QLabel("Width:"))
        
        self.column_width_spin = QDoubleSpinBox()
        self.column_width_spin.setRange(0.1, 22.0)  # 0.1 to 22 inches
        self.column_width_spin.setValue(1.0)
        self.column_width_spin.setSuffix(" \"")
        
        self.column_width_combo = QComboBox()
        self.column_width_combo.addItems(["Inches", "Centimeters", "Percent"])
        
        size_layout.addWidget(self.column_width_spin)
        size_layout.addWidget(self.column_width_combo)
        size_group.setLayout(size_layout)
        
        # Column width controls
        width_controls = QHBoxLayout()
        
        self.prev_column_btn = QPushButton("Previous Column")
        self.next_column_btn = QPushButton("Next Column")
        
        width_controls.addWidget(self.prev_column_btn)
        width_controls.addWidget(self.next_column_btn)
        
        # AutoFit behavior
        autofit_group = QGroupBox("AutoFit Behavior")
        autofit_layout = QVBoxLayout()
        
        self.autofit_none_radio = QRadioButton("Fixed column width")
        self.autofit_contents_radio = QRadioButton("AutoFit to contents")
        self.autofit_window_radio = QRadioButton("AutoFit to window")
        
        self.autofit_none_radio.setChecked(True)
        
        autofit_layout.addWidget(self.autofit_none_radio)
        autofit_layout.addWidget(self.autofit_contents_radio)
        autofit_layout.addWidget(self.autofit_window_radio)
        autofit_group.setLayout(autofit_layout)
        
        # Add groups to layout
        layout.addWidget(size_group)
        layout.addLayout(width_controls)
        layout.addWidget(autofit_group)
        layout.addStretch()
    
    def setup_cell_tab(self, parent):
        """Setup the Cell tab."""
        layout = QVBoxLayout(parent)
        
        # Size group
        size_group = QGroupBox("Size")
        size_layout = QHBoxLayout()
        
        size_layout.addWidget(QLabel("Width:"))
        
        self.cell_width_spin = QDoubleSpinBox()
        self.cell_width_spin.setRange(0.1, 22.0)  # 0.1 to 22 inches
        self.cell_width_spin.setValue(1.0)
        self.cell_width_spin.setSuffix(" \"")
        
        size_layout.addWidget(self.cell_width_spin)
        size_group.setLayout(size_layout)
        
        # Options
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()
        
        self.wrap_text_check = QCheckBox("Wrap text")
        self.wrap_text_check.setChecked(True)
        
        self.fit_text_check = QCheckBox("Fit text")
        
        options_layout.addWidget(self.wrap_text_check)
        options_layout.addWidget(self.fit_text_check)
        options_layout.addStretch()
        options_group.setLayout(options_layout)
        
        # Margins
        margins_group = QGroupBox("Cell Margins")
        margins_layout = QGridLayout()
        
        self.margin_top_spin = QDoubleSpinBox()
        self.margin_top_spin.setRange(0, 31680)  # 22 inches in points
        self.margin_top_spin.setValue(self.properties['cell_margins']['top'])
        self.margin_top_spin.setSuffix(" pt")
        
        self.margin_bottom_spin = QDoubleSpinBox()
        self.margin_bottom_spin.setRange(0, 31680)
        self.margin_bottom_spin.setValue(self.properties['cell_margins']['bottom'])
        self.margin_bottom_spin.setSuffix(" pt")
        
        self.margin_left_spin = QDoubleSpinBox()
        self.margin_left_spin.setRange(0, 31680)
        self.margin_left_spin.setValue(self.properties['cell_margins']['left'])
        self.margin_left_spin.setSuffix(" pt")
        
        self.margin_right_spin = QDoubleSpinBox()
        self.margin_right_spin.setRange(0, 31680)
        self.margin_right_spin.setValue(self.properties['cell_margins']['right'])
        self.margin_right_spin.setSuffix(" pt")
        
        margins_layout.addWidget(QLabel("Top:"), 0, 0)
        margins_layout.addWidget(self.margin_top_spin, 0, 1)
        margins_layout.addWidget(QLabel("Bottom:"), 1, 0)
        margins_layout.addWidget(self.margin_bottom_spin, 1, 1)
        margins_layout.addWidget(QLabel("Left:"), 2, 0)
        margins_layout.addWidget(self.margin_left_spin, 2, 1)
        margins_layout.addWidget(QLabel("Right:"), 3, 0)
        margins_layout.addWidget(self.margin_right_spin, 3, 1)
        
        # Default margins button
        default_btn = QPushButton("Default...")
        default_btn.clicked.connect(self.reset_margins)
        
        margins_layout.addWidget(default_btn, 4, 0, 1, 2)
        margins_group.setLayout(margins_layout)
        
        # Add groups to layout
        layout.addWidget(size_group)
        layout.addWidget(options_group)
        layout.addWidget(margins_group)
        layout.addStretch()
    
    def setup_borders_tab(self, parent):
        """Setup the Borders and Shading tab."""
        self.borders_tabs = QTabWidget()
        
        # Borders tab
        borders_tab = QWidget()
        borders_layout = QVBoxLayout(borders_tab)
        
        # Settings group
        settings_group = QGroupBox("Settings")
        settings_layout = QHBoxLayout()
        
        # Border style list
        self.border_style_list = QListWidget()
        self.border_style_list.addItems(["None", "Single", "Double", "Dashed", "Dotted", "Wavy"])
        self.border_style_list.setCurrentRow(1)  # Select Single by default
        self.border_style_list.currentRowChanged.connect(self.update_border_style)
        
        # Border color button
        self.border_color_btn = QPushButton()
        self.border_color_btn.setFixedSize(24, 24)
        self.border_color_btn.setStyleSheet(f"background-color: {self.properties['borders']['all']['color'].name()}; border: 1px solid black;")
        self.border_color_btn.clicked.connect(self.choose_border_color)
        
        # Border width
        self.border_width_combo = QComboBox()
        self.border_width_combo.addItems(["0.5 pt", "1 pt", "1.5 pt", "2.25 pt", "3 pt", "4.5 pt", "6 pt"])
        self.border_width_combo.setCurrentText(f"{self.properties['borders']['all']['width']} pt")
        self.border_width_combo.currentTextChanged.connect(self.update_border_width)
        
        # Preview area
        self.border_preview = QWidget()
        self.border_preview.setMinimumSize(150, 150)
        self.border_preview.setStyleSheet("background-color: white; border: 1px solid #ccc;")
        
        # Add to settings layout
        settings_layout.addWidget(QLabel("Style:"))
        settings_layout.addWidget(self.border_style_list, 1)
        settings_layout.addWidget(QLabel("Color:"))
        settings_layout.addWidget(self.border_color_btn)
        settings_layout.addWidget(QLabel("Width:"))
        settings_layout.addWidget(self.border_width_combo)
        
        settings_group.setLayout(settings_layout)
        
        # Preview group
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()
        preview_layout.addWidget(self.border_preview, 1, Qt.AlignCenter)
        preview_group.setLayout(preview_layout)
        
        # Buttons for applying to specific borders
        border_btns_layout = QGridLayout()
        
        border_btns = [
            ("Top", 0, 1), ("Bottom", 2, 1), ("Left", 1, 0), ("Right", 1, 2),
            ("Inside H", 3, 1), ("Inside V", 1, 3), ("All", 1, 1)
        ]
        
        for text, row, col in border_btns:
            btn = QPushButton(text)
            btn.setCheckable(True)
            btn.setChecked(True)
            border_btns_layout.addWidget(btn, row, col)
        
        # Add to borders layout
        borders_layout.addWidget(settings_group)
        borders_layout.addLayout(border_btns_layout)
        borders_layout.addWidget(preview_group)
        
        # Shading tab
        shading_tab = QWidget()
        shading_layout = QVBoxLayout(shading_tab)
        
        # Fill color
        fill_group = QGroupBox("Fill")
        fill_layout = QHBoxLayout()
        
        self.fill_color_btn = QPushButton()
        self.fill_color_btn.setFixedSize(24, 24)
        self.fill_color_btn.setStyleSheet(f"background-color: {self.properties['shading']['fill'].name()}; border: 1px solid black;")
        self.fill_color_btn.clicked.connect(self.choose_fill_color)
        
        self.fill_none_btn = QPushButton("No Color")
        self.fill_none_btn.clicked.connect(lambda: self.set_fill_color(Qt.transparent))
        
        fill_layout.addWidget(QLabel("Color:"))
        fill_layout.addWidget(self.fill_color_btn)
        fill_layout.addWidget(self.fill_none_btn)
        fill_layout.addStretch()
        fill_group.setLayout(fill_layout)
        
        # Pattern
        pattern_group = QGroupBox("Pattern")
        pattern_layout = QGridLayout()
        
        patterns = [
            "5%", "10%", "12.5%", "20%", "25%", "30%", "37.5%", "40%",
            "50%", "60%", "62.5%", "70%", "75%", "80%", "87.5%", "90%"
        ]
        
        for i, pattern in enumerate(patterns):
            btn = QPushButton()
            btn.setFixedSize(30, 30)
            btn.setStyleSheet(f"background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 black, stop:{pattern} black, stop:{pattern} white, stop:1 white);")
            pattern_layout.addWidget(btn, i // 4, i % 4)
        
        pattern_group.setLayout(pattern_layout)
        
        # Add to shading layout
        shading_layout.addWidget(fill_group)
        shading_layout.addWidget(pattern_group)
        shading_layout.addStretch()
        
        # Add tabs
        self.borders_tabs.addTab(borders_tab, "Borders")
        self.borders_tabs.addTab(shading_tab, "Shading")
        
        # Set layout
        parent_layout = QVBoxLayout(parent)
        parent_layout.addWidget(self.borders_tabs)
    
    def setup_alt_text_tab(self, parent):
        """Setup the Alt Text tab."""
        layout = QVBoxLayout(parent)
        
        # Description
        desc_group = QGroupBox("Description")
        desc_layout = QVBoxLayout()
        
        self.desc_edit = QLineEdit()
        self.desc_edit.setPlaceholderText("Description (required)")
        desc_layout.addWidget(self.desc_edit)
        
        desc_help = QLabel("Describe the content and context of the table for users with screen readers.")
        desc_help.setWordWrap(True)
        desc_help.setStyleSheet("color: #666; font-style: italic;")
        desc_layout.addWidget(desc_help)
        
        desc_group.setLayout(desc_layout)
        
        # Title
        title_group = QGroupBox("Title (optional)")
        title_layout = QVBoxLayout()
        
        self.title_edit = QLineEdit()
        title_layout.addWidget(self.title_edit)
        
        title_help = QLabel("A short summary of the table contents.")
        title_help.setWordWrap(True)
        title_help.setStyleSheet("color: #666; font-style: italic;")
        title_layout.addWidget(title_help)
        
        title_group.setLayout(title_layout)
        
        # Add groups to layout
        layout.addWidget(desc_group)
        layout.addWidget(title_group)
        layout.addStretch()
    
    def update_preview(self):
        """Update the table preview."""
        # This would be implemented to show a live preview of the table
        # with the current settings applied
        pass
    
    def update_table_width(self, value):
        """Update the table width."""
        self.properties['preferred_width']['value'] = value
        self.update_preview()
    
    def update_table_width_unit(self, unit):
        """Update the table width unit."""
        self.properties['preferred_width']['unit'] = unit.lower()
        self.update_preview()
    
    def update_alignment(self, alignment):
        """Update the table alignment."""
        self.properties['alignment'] = alignment
        self.update_preview()
    
    def update_text_wrapping(self, checked):
        """Update text wrapping setting."""
        self.properties['text_wrapping'] = not checked  # None is checked when no wrapping
        self.positioning_group.setEnabled(self.properties['text_wrapping'])
        self.update_preview()
    
    def update_break_across_pages(self, checked):
        """Update allow break across pages setting."""
        self.properties['allow_break_across_pages'] = checked
    
    def update_repeat_header(self, checked):
        """Update repeat header row setting."""
        self.properties['repeat_header_row'] = checked
    
    def reset_margins(self):
        """Reset cell margins to default values."""
        defaults = {
            'top': 0.0,
            'bottom': 0.0,
            'left': 0.08,  # 0.08 inches = 5.76 points
            'right': 0.08
        }
        
        self.properties['cell_margins'].update(defaults)
        
        self.margin_top_spin.setValue(defaults['top'])
        self.margin_bottom_spin.setValue(defaults['bottom'])
        self.margin_left_spin.setValue(defaults['left'])
        self.margin_right_spin.setValue(defaults['right'])
    
    def update_border_style(self, index):
        """Update the border style."""
        style_map = {
            0: 'none', 1: 'solid', 2: 'double', 
            3: 'dashed', 4: 'dotted', 5: 'wavy'
        }
        
        style = style_map.get(index, 'solid')
        
        # Update all borders with the new style
        for border in self.properties['borders'].values():
            border['style'] = style
        
        self.update_preview()
    
    def update_border_width(self, text):
        """Update the border width."""
        try:
            width = float(text.split()[0])  # Extract the number from "X pt"
            
            # Update all borders with the new width
            for border in self.properties['borders'].values():
                border['width'] = width
            
            self.update_preview()
        except (ValueError, IndexError):
            pass
    
    def choose_border_color(self):
        """Open a color dialog to choose the border color."""
        color = QColorDialog.getColor(self.properties['borders']['all']['color'], self, "Border Color")
        if color.isValid():
            # Update all borders with the new color
            for border in self.properties['borders'].values():
                border['color'] = color
            
            self.border_color_btn.setStyleSheet(f"background-color: {color.name()}; border: 1px solid black;")
            self.update_preview()
    
    def choose_fill_color(self):
        """Open a color dialog to choose the fill color."""
        color = QColorDialog.getColor(self.properties['shading']['fill'], self, "Fill Color")
        if color.isValid():
            self.set_fill_color(color)
    
    def set_fill_color(self, color):
        """Set the fill color."""
        self.properties['shading']['fill'] = color
        
        if color == Qt.transparent:
            self.fill_color_btn.setStyleSheet("background-color: white; border: 1px solid #ccc;")
        else:
            self.fill_color_btn.setStyleSheet(f"background-color: {color.name()}; border: 1px solid black;")
        
        self.update_preview()
    
    def get_values(self):
        """Get the current table properties as a dictionary."""
        # Update properties from UI elements
        self.properties['cell_margins'].update({
            'top': self.margin_top_spin.value(),
            'bottom': self.margin_bottom_spin.value(),
            'left': self.margin_left_spin.value(),
            'right': self.margin_right_spin.value()
        })
        
        # Update description if in alt text tab
        if self.tabs.currentWidget() == self.tabs.widget(4):  # Alt Text tab
            self.properties['description'] = self.desc_edit.text()
        
        return self.properties
    
    def set_values(self, values):
        """Set the table properties from a dictionary."""
        if not values:
            return
        
        self.properties.update(values)
        
        # Update UI elements
        # Table tab
        self.width_spin.setValue(self.properties['preferred_width']['value'])
        self.width_unit_combo.setCurrentText(self.properties['preferred_width']['unit'].capitalize())
        
        for btn in self.align_buttons.buttons():
            if self.align_buttons.id(btn) == self.properties['alignment']:
                btn.setChecked(True)
                break
        
        self.wrap_none_radio.setChecked(not self.properties['text_wrapping'])
        self.wrap_around_radio.setChecked(self.properties['text_wrapping'])
        self.positioning_group.setEnabled(self.properties['text_wrapping'])
        
        # Cell tab
        if 'cell_margins' in self.properties:
            margins = self.properties['cell_margins']
            self.margin_top_spin.setValue(margins.get('top', 0.0))
            self.margin_bottom_spin.setValue(margins.get('bottom', 0.0))
            self.margin_left_spin.setValue(margins.get('left', 0.1))
            self.margin_right_spin.setValue(margins.get('right', 0.1))
        
        # Update preview
        self.update_preview()
