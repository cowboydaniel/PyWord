from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QDoubleSpinBox, QGroupBox, QRadioButton, QButtonGroup,
                             QDialogButtonBox, QCheckBox, QTabWidget, QSizePolicy)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QPageLayout, QPageSize, QPagedPaintDevice

class PageSetupDialog(QTabWidget):
    """Page setup dialog with tabs for page size, margins, layout, and paper source."""
    
    # Paper sizes (name, QPageSize enum, width, height in mm)
    PAPER_SIZES = [
        ("Letter", QPageSize.Letter, 215.9, 279.4),
        ("Legal", QPageSize.Legal, 215.9, 355.6),
        ("A4", QPageSize.A4, 210, 297),
        ("A5", QPageSize.A5, 148, 210),
        ("A3", QPageSize.A3, 297, 420),
        ("Tabloid", QPageSize.Tabloid, 279.4, 431.8),
        ("Executive", QPageSize.Executive, 184.15, 266.7),
        ("Custom", None, 0, 0)
    ]
    
    # Units for measurement
    UNITS = ["Millimeters", "Inches"]
    
    def __init__(self, parent=None, page_layout=None):
        """Initialize the page setup dialog.
        
        Args:
            parent: Parent widget
            page_layout: Optional QPageLayout with initial settings
        """
        super().__init__(parent)
        self.setWindowTitle("Page Setup")
        self.setMinimumSize(500, 400)
        
        # Default page layout
        self.page_layout = page_layout if page_layout else QPageLayout()
        
        # Current units (0=mm, 1=inches)
        self.current_units = 0
        
        # Setup UI
        self.setup_ui()
        
        # Update UI with initial values
        self.update_ui_from_page_layout()
    
    def setup_ui(self):
        """Setup the page setup dialog UI."""
        # Create tabs
        self.setup_margins_tab()
        self.setup_paper_tab()
        self.setup_layout_tab()
        
        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Apply | QDialogButtonBox.Help
        )
        
        # Connect signals
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        apply_btn = self.button_box.button(QDialogButtonBox.Apply)
        apply_btn.clicked.connect(self.apply_changes)
        
        # Main layout
        main_layout = QVBoxLayout()
        main_layout.addWidget(self)
        main_layout.addWidget(self.button_box)
        self.setLayout(main_layout)
    
    def setup_margins_tab(self):
        """Setup the Margins tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Units
        units_group = QGroupBox("Units")
        units_layout = QHBoxLayout()
        
        self.units_combo = QComboBox()
        self.units_combo.addItems(self.UNITS)
        self.units_combo.currentIndexChanged.connect(self.on_units_changed)
        
        units_layout.addWidget(QLabel("Units:"))
        units_layout.addWidget(self.units_combo)
        units_layout.addStretch()
        units_group.setLayout(units_layout)
        
        # Margins
        margins_group = QGroupBox("Margins")
        margins_layout = QGridLayout()
        
        # Top margin
        self.top_margin = QDoubleSpinBox()
        self.top_margin.setRange(0, 1000)
        self.top_margin.setSuffix(" mm")
        self.top_margin.setDecimals(1)
        self.top_margin.setSingleStep(1.0)
        
        # Bottom margin
        self.bottom_margin = QDoubleSpinBox()
        self.bottom_margin.setRange(0, 1000)
        self.bottom_margin.setSuffix(" mm")
        self.bottom_margin.setDecimals(1)
        self.bottom_margin.setSingleStep(1.0)
        
        # Left margin
        self.left_margin = QDoubleSpinBox()
        self.left_margin.setRange(0, 1000)
        self.left_margin.setSuffix(" mm")
        self.left_margin.setDecimals(1)
        self.left_margin.setSingleStep(1.0)
        
        # Right margin
        self.right_margin = QDoubleSpinBox()
        self.right_margin.setRange(0, 1000)
        self.right_margin.setSuffix(" mm")
        self.right_margin.setDecimals(1)
        self.right_margin.setSingleStep(1.0)
        
        # Gutter
        self.gutter = QDoubleSpinBox()
        self.gutter.setRange(0, 1000)
        self.gutter.setSuffix(" mm")
        self.gutter.setDecimals(1)
        self.gutter.setSingleStep(1.0)
        
        # Add to layout
        margins_layout.addWidget(QLabel("Top:"), 0, 0)
        margins_layout.addWidget(self.top_margin, 0, 1)
        margins_layout.addWidget(QLabel("Bottom:"), 1, 0)
        margins_layout.addWidget(self.bottom_margin, 1, 1)
        margins_layout.addWidget(QLabel("Left:"), 2, 0)
        margins_layout.addWidget(self.left_margin, 2, 1)
        margins_layout.addWidget(QLabel("Right:"), 3, 0)
        margins_layout.addWidget(self.right_margin, 3, 1)
        margins_layout.addWidget(QLabel("Gutter:"), 4, 0)
        margins_layout.addWidget(self.gutter, 4, 1)
        
        # Preview
        self.margins_preview = QLabel()
        self.margins_preview.setAlignment(Qt.AlignCenter)
        self.margins_preview.setMinimumSize(200, 200)
        self.margins_preview.setStyleSheet("background-color: white; border: 1px solid #ccc;")
        
        # Add to margins layout
        margins_layout.addWidget(self.margins_preview, 0, 2, 5, 1)
        
        margins_group.setLayout(margins_layout)
        
        # Add to tab
        layout.addWidget(units_group)
        layout.addWidget(margins_group)
        layout.addStretch()
        
        # Add tab
        self.addTab(tab, "Margins")
    
    def setup_paper_tab(self):
        """Setup the Paper tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Paper size
        size_group = QGroupBox("Paper Size")
        size_layout = QVBoxLayout()
        
        # Paper size combo
        self.paper_size_combo = QComboBox()
        for size in self.PAPER_SIZES:
            self.paper_size_combo.addItem(size[0])
        
        # Connect signal after adding items
        self.paper_size_combo.currentIndexChanged.connect(self.on_paper_size_changed)
        
        # Width and height
        self.paper_width = QDoubleSpinBox()
        self.paper_width.setRange(1, 1000)
        self.paper_width.setSuffix(" mm")
        self.paper_width.setDecimals(1)
        self.paper_width.valueChanged.connect(self.on_paper_dimensions_changed)
        
        self.paper_height = QDoubleSpinBox()
        self.paper_height.setRange(1, 1000)
        self.paper_height.setSuffix(" mm")
        self.paper_height.setDecimals(1)
        self.paper_height.valueChanged.connect(self.on_paper_dimensions_changed)
        
        # Orientation
        orientation_group = QGroupBox("Orientation")
        orientation_layout = QHBoxLayout()
        
        self.portrait_radio = QRadioButton("Portrait")
        self.landscape_radio = QRadioButton("Landscape")
        
        self.portrait_radio.toggled.connect(self.on_orientation_changed)
        
        orientation_layout.addWidget(self.portrait_radio)
        orientation_layout.addWidget(self.landscape_radio)
        orientation_layout.addStretch()
        orientation_group.setLayout(orientation_layout)
        
        # Add to size layout
        size_layout.addWidget(QLabel("Paper size:"))
        size_layout.addWidget(self.paper_size_combo)
        size_layout.addSpacing(10)
        size_layout.addWidget(QLabel("Width:"))
        size_layout.addWidget(self.paper_width)
        size_layout.addWidget(QLabel("Height:"))
        size_layout.addWidget(self.paper_height)
        size_layout.addSpacing(10)
        size_layout.addWidget(orientation_group)
        size_layout.addStretch()
        
        size_group.setLayout(size_layout)
        
        # Paper source
        source_group = QGroupBox("Paper Source")
        source_layout = QVBoxLayout()
        
        # First page
        first_page_layout = QHBoxLayout()
        first_page_layout.addWidget(QLabel("First page:"))
        
        self.first_page_combo = QComboBox()
        self.first_page_combo.addItems(["Default tray"])
        
        first_page_layout.addWidget(self.first_page_combo)
        first_page_layout.addStretch()
        
        # Other pages
        other_pages_layout = QHBoxLayout()
        other_pages_layout.addWidget(QLabel("Other pages:"))
        
        self.other_pages_combo = QComboBox()
        self.other_pages_combo.addItems(["Same as first page"])
        
        other_pages_layout.addWidget(self.other_pages_combo)
        other_pages_layout.addStretch()
        
        # Add to source layout
        source_layout.addLayout(first_page_layout)
        source_layout.addLayout(other_pages_layout)
        source_group.setLayout(source_layout)
        
        # Add to tab
        layout.addWidget(size_group)
        layout.addWidget(source_group)
        layout.addStretch()
        
        # Add tab
        self.addTab(tab, "Paper")
    
    def setup_layout_tab(self):
        """Setup the Layout tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Section start
        section_group = QGroupBox("Section")
        section_layout = QVBoxLayout()
        
        self.section_start_combo = QComboBox()
        self.section_start_combo.addItems([
            "New page",
            "Continuous",
            "Even page",
            "Odd page"
        ])
        
        section_layout.addWidget(QLabel("Section start:"))
        section_layout.addWidget(self.section_start_combo)
        section_layout.addStretch()
        section_group.setLayout(section_layout)
        
        # Headers and footers
        headers_group = QGroupBox("Headers and Footers")
        headers_layout = QVBoxLayout()
        
        self.different_odd_even = QCheckBox("Different odd and even pages")
        self.different_first = QCheckBox("Different first page")
        
        # Vertical alignment
        align_group = QGroupBox("Vertical Alignment")
        align_layout = QVBoxLayout()
        
        self.align_top = QRadioButton("Top")
        self.align_center = QRadioButton("Center")
        self.align_justified = QRadioButton("Justified")
        self.align_bottom = QRadioButton("Bottom")
        
        align_layout.addWidget(self.align_top)
        align_layout.addWidget(self.align_center)
        align_layout.addWidget(self.align_justified)
        align_layout.addWidget(self.align_bottom)
        align_group.setLayout(align_layout)
        
        # Line numbers
        line_nums_group = QGroupBox("Line Numbers")
        line_nums_layout = QVBoxLayout()
        
        self.line_nums_check = QCheckBox("Add line numbering")
        self.line_nums_button = QPushButton("Line Numbers...")
        self.line_nums_button.setEnabled(False)
        
        self.line_nums_check.toggled.connect(self.line_nums_button.setEnabled)
        
        line_nums_layout.addWidget(self.line_nums_check)
        line_nums_layout.addWidget(self.line_nums_button, 0, Qt.AlignLeft)
        line_nums_layout.addStretch()
        line_nums_group.setLayout(line_nums_layout)
        
        # Add to headers layout
        headers_layout.addWidget(self.different_odd_even)
        headers_layout.addWidget(self.different_first)
        headers_layout.addWidget(align_group)
        headers_layout.addWidget(line_nums_group)
        headers_group.setLayout(headers_layout)
        
        # Preview
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()
        
        self.preview = QLabel()
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setMinimumSize(200, 200)
        self.preview.setStyleSheet("background-color: white; border: 1px solid #ccc;")
        
        preview_layout.addWidget(self.preview)
        preview_group.setLayout(preview_layout)
        
        # Add to tab
        layout.addWidget(section_group)
        layout.addWidget(headers_group)
        layout.addWidget(preview_group)
        layout.addStretch()
        
        # Add tab
        self.addTab(tab, "Layout")
    
    def update_ui_from_page_layout(self):
        """Update the UI with values from the current page layout."""
        # Margins
        margins = self.page_layout.marginsPoints()
        dpi = 72  # Standard screen DPI for points to inches conversion
        mm_per_inch = 25.4
        
        # Convert points to mm
        self.top_margin.setValue(margins.top() * mm_per_inch / dpi)
        self.bottom_margin.setValue(margins.bottom() * mm_per_inch / dpi)
        self.left_margin.setValue(margins.left() * mm_per_inch / dpi)
        self.right_margin.setValue(margins.right() * mm_per_inch / dpi)
        
        # Paper size
        page_size = self.page_layout.pageSize()
        size_mm = page_size.size(QPageSize.Millimeter)
        
        self.paper_width.setValue(size_mm.width())
        self.paper_height.setValue(size_mm.height())
        
        # Orientation
        if self.page_layout.orientation() == QPageLayout.Portrait:
            self.portrait_radio.setChecked(True)
        else:
            self.landscape_radio.setChecked(True)
        
        # Update paper size combo
        self.update_paper_size_combo()
        
        # Update preview
        self.update_preview()
    
    def update_page_layout_from_ui(self):
        """Update the page layout with values from the UI."""
        # Convert mm to points (1 inch = 25.4 mm = 72 points)
        mm_per_inch = 25.4
        points_per_inch = 72
        mm_to_points = points_per_inch / mm_per_inch
        
        # Update margins
        top = self.top_margin.value() * mm_to_points
        bottom = self.bottom_margin.value() * mm_to_points
        left = self.left_margin.value() * mm_to_points
        right = self.right_margin.value() * mm_to_points
        
        self.page_layout.setMargins(QMarginsF(left, top, right, bottom))
        
        # Update paper size
        width_mm = self.paper_width.value()
        height_mm = self.paper_height.value()
        
        # Find matching paper size
        size_found = False
        for i, (name, qsize, width, height) in enumerate(self.PAPER_SIZES):
            if qsize is not None:
                size = QPageSize(qsize)
                size_mm = size.size(QPageSize.Millimeter)
                
                # Check if dimensions match (with small tolerance for floating point)
                if (abs(size_mm.width() - width_mm) < 1.0 and 
                    abs(size_mm.height() - height_mm) < 1.0):
                    self.page_layout.setPageSize(size)
                    size_found = True
                    break
        
        # If no matching standard size, create a custom size
        if not size_found:
            custom_size = QPageSize(QSizeF(width_mm, height_mm), QPageSize.Millimeter)
            self.page_layout.setPageSize(custom_size)
        
        # Update orientation
        if self.portrait_radio.isChecked():
            self.page_layout.setOrientation(QPageLayout.Portrait)
        else:
            self.page_layout.setOrientation(QPageLayout.Landscape)
    
    def update_paper_size_combo(self):
        """Update the paper size combo box based on current dimensions."""
        width_mm = self.paper_width.value()
        height_mm = self.paper_height.value()
        
        # Find matching paper size
        found_index = len(self.PAPER_SIZES) - 1  # Default to Custom
        
        for i, (name, qsize, width, height) in enumerate(self.PAPER_SIZES):
            if qsize is not None:
                # Check if dimensions match (with small tolerance for floating point)
                if (abs(width - width_mm) < 1.0 and abs(height - height_mm) < 1.0) or \
                   (abs(width - height_mm) < 1.0 and abs(height - width_mm) < 1.0):
                    found_index = i
                    break
        
        # Block signals to avoid triggering on_paper_size_changed
        self.paper_size_combo.blockSignals(True)
        self.paper_size_combo.setCurrentIndex(found_index)
        self.paper_size_combo.blockSignals(False)
    
    def update_preview(self):
        """Update the preview with current settings."""
        # This would be implemented to show a preview of the page with margins
        pass
    
    def on_units_changed(self, index):
        """Handle units change (mm/inches)."""
        # Convert values if needed
        if index != self.current_units:
            # Convert between mm and inches (1 inch = 25.4 mm)
            factor = 25.4 if index == 1 else 1/25.4
            
            # Update margin spinboxes
            self.top_margin.setValue(self.top_margin.value() * factor)
            self.bottom_margin.setValue(self.bottom_margin.value() * factor)
            self.left_margin.setValue(self.left_margin.value() * factor)
            self.right_margin.setValue(self.right_margin.value() * factor)
            self.gutter.setValue(self.gutter.value() * factor)
            
            # Update paper size spinboxes
            self.paper_width.setValue(self.paper_width.value() * factor)
            self.paper_height.setValue(self.paper_height.value() * factor)
            
            # Update suffixes
            suffix = " in" if index == 1 else " mm"
            
            for spinbox in [self.top_margin, self.bottom_margin, self.left_margin, 
                           self.right_margin, self.gutter, self.paper_width, self.paper_height]:
                spinbox.setSuffix(suffix)
            
            # Update current units
            self.current_units = index
    
    def on_paper_size_changed(self, index):
        """Handle paper size change."""
        if 0 <= index < len(self.PAPER_SIZES):
            name, qsize, width, height = self.PAPER_SIZES[index]
            
            if qsize is not None:  # Standard size
                # Get size in current units
                if self.current_units == 0:  # mm
                    self.paper_width.setValue(width)
                    self.paper_height.setValue(height)
                else:  # inches
                    self.paper_width.setValue(round(width / 25.4, 2))
                    self.paper_height.setValue(round(height / 25.4, 2))
                
                # Disable width/height editing for standard sizes
                self.paper_width.setEnabled(False)
                self.paper_height.setEnabled(False)
            else:  # Custom size
                # Enable width/height editing
                self.paper_width.setEnabled(True)
                self.paper_height.setEnabled(True)
    
    def on_paper_dimensions_changed(self, value):
        """Handle paper dimension changes."""
        # If custom size is selected or we're changing to a custom size
        if self.paper_size_combo.currentIndex() == len(self.PAPER_SIZES) - 1:
            # Update paper size combo to Custom
            self.paper_size_combo.blockSignals(True)
            self.paper_size_combo.setCurrentIndex(len(self.PAPER_SIZES) - 1)
            self.paper_size_combo.blockSignals(False)
            
            # Enable width/height editing
            self.paper_width.setEnabled(True)
            self.paper_height.setEnabled(True)
        else:
            # Check if dimensions match any standard size
            self.update_paper_size_combo()
    
    def on_orientation_changed(self, portrait):
        """Handle orientation change."""
        # Swap width and height if needed
        if (portrait and self.paper_width.value() > self.paper_height.value()) or \
           (not portrait and self.paper_width.value() < self.paper_height.value()):
            width = self.paper_width.value()
            height = self.paper_height.value()
            
            self.paper_width.blockSignals(True)
            self.paper_height.blockSignals(True)
            
            self.paper_width.setValue(height)
            self.paper_height.setValue(width)
            
            self.paper_width.blockSignals(False)
            self.paper_height.blockSignals(False)
    
    def apply_changes(self):
        """Apply changes to the page layout."""
        self.update_page_layout_from_ui()
    
    def accept(self):
        """Handle dialog acceptance."""
        self.apply_changes()
        super().accept()
    
    def get_page_layout(self):
        """Get the current page layout."""
        return self.page_layout
