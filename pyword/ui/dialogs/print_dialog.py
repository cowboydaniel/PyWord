from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QSpinBox, QGroupBox, QRadioButton, QButtonGroup, QCheckBox,
                             QDialog, QDialogButtonBox, QPushButton, QTabWidget, QListWidget,
                             QListWidgetItem, QLineEdit, QFormLayout, QGridLayout, QSpacerItem,
                             QSizePolicy)
from PySide6.QtCore import Qt, QSize, QMarginsF
from PySide6.QtGui import QPageLayout, QPageSize, QPagedPaintDevice, QPainter, QTextDocument
from PySide6.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog

class PrintDialog(QDialog):
    """Print dialog with print preview and options."""
    
    def __init__(self, document, parent=None):
        """Initialize the print dialog.
        
        Args:
            document: QTextDocument to print
            parent: Parent widget
        """
        super().__init__(parent)
        self.setWindowTitle("Print")
        self.setMinimumSize(700, 600)
        
        self.document = document
        self.printer = QPrinter(QPrinter.HighResolution)
        self.printer.setPageMargins(QMarginsF(20, 20, 20, 20), QPageLayout.Millimeter)
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the print dialog UI."""
        main_layout = QHBoxLayout(self)
        
        # Left side: Print options
        options_widget = QWidget()
        options_layout = QVBoxLayout(options_widget)
        
        # Printer selection
        printer_group = QGroupBox("Printer")
        printer_layout = QVBoxLayout()
        
        # Printer name and status
        self.printer_name = QLabel(self.printer.printerName())
        self.printer_status = QLabel("Ready")
        self.printer_status.setStyleSheet("color: green;")
        
        # Printer properties button
        self.properties_btn = QPushButton("Properties...")
        self.properties_btn.clicked.connect(self.show_printer_properties)
        
        # Find printer button
        self.find_printer_btn = QPushButton("Find Printer...")
        
        printer_layout.addWidget(QLabel("Name:"))
        printer_layout.addWidget(self.printer_name)
        printer_layout.addWidget(QLabel("Status:"))
        printer_layout.addWidget(self.printer_status)
        
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.properties_btn)
        btn_layout.addWidget(self.find_printer_btn)
        
        printer_layout.addLayout(btn_layout)
        printer_group.setLayout(printer_layout)
        
        # Page range
        range_group = QGroupBox("Page Range")
        range_layout = QVBoxLayout()
        
        self.range_all = QRadioButton("All")
        self.range_current = QRadioButton("Current page")
        self.range_selection = QRadioButton("Selection")
        self.range_custom = QRadioButton("Pages:")
        
        self.range_edit = QLineEdit()
        self.range_edit.setEnabled(False)
        self.range_edit.setPlaceholderText("e.g., 1-5, 8, 11-13")
        
        self.range_all.setChecked(True)
        self.range_custom.toggled.connect(self.range_edit.setEnabled)
        
        range_layout.addWidget(self.range_all)
        range_layout.addWidget(self.range_current)
        range_layout.addWidget(self.range_selection)
        
        custom_layout = QHBoxLayout()
        custom_layout.addWidget(self.range_custom)
        custom_layout.addWidget(self.range_edit)
        
        range_layout.addLayout(custom_layout)
        
        # Page range example
        example = QLabel("Example: 1-5, 8, 11-13")
        example.setStyleSheet("color: #666; font-style: italic;")
        
        range_layout.addWidget(example)
        range_group.setLayout(range_layout)
        
        # Copies
        copies_group = QGroupBox("Copies")
        copies_layout = QGridLayout()
        
        self.copies_spin = QSpinBox()
        self.copies_spin.setRange(1, 999)
        self.copies_spin.setValue(1)
        
        self.collate_check = QCheckBox("Collate")
        self.collate_check.setChecked(True)
        
        self.print_to_file = QCheckBox("Print to file")
        
        copies_layout.addWidget(QLabel("Number of copies:"), 0, 0)
        copies_layout.addWidget(self.copies_spin, 0, 1)
        copies_layout.addWidget(self.collate_check, 1, 0, 1, 2)
        copies_layout.addWidget(self.print_to_file, 2, 0, 1, 2)
        
        copies_group.setLayout(copies_layout)
        
        # Print what
        print_what_group = QGroupBox("Print What")
        print_what_layout = QVBoxLayout()
        
        self.print_what_combo = QComboBox()
        self.print_what_combo.addItems(["Document", "Document with markup", "List of markup"])
        
        print_what_layout.addWidget(self.print_what_combo)
        print_what_group.setLayout(print_what_layout)
        
        # Options
        options_group = QGroupBox("Options")
        options_inner_layout = QVBoxLayout()
        
        self.print_hidden_text = QCheckBox("Print hidden text")
        self.print_background = QCheckBox("Print background colors and images")
        self.update_fields = QCheckBox("Update fields before printing")
        self.update_links = QCheckBox("Update links before printing")
        
        options_inner_layout.addWidget(self.print_hidden_text)
        options_inner_layout.addWidget(self.print_background)
        options_inner_layout.addWidget(self.update_fields)
        options_inner_layout.addWidget(self.update_links)
        options_group.setLayout(options_inner_layout)
        
        # Add all groups to options layout
        options_layout.addWidget(printer_group)
        options_layout.addWidget(range_group)
        options_layout.addWidget(copies_group)
        options_layout.addWidget(print_what_group)
        options_layout.addWidget(options_group)
        options_layout.addStretch()
        
        # Right side: Preview
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()
        
        # Preview controls
        controls_layout = QHBoxLayout()
        
        self.page_label = QLabel("Page 1 of 1")
        self.zoom_combo = QComboBox()
        self.zoom_combo.addItems(["500%", "200%", "150%", "100%", "75%", "50%", "25%", "10%", "Page Width", "Text Width", "Whole Page"])
        self.zoom_combo.setCurrentText("100%")
        
        controls_layout.addWidget(QLabel("Zoom:"))
        controls_layout.addWidget(self.zoom_combo)
        controls_layout.addStretch()
        controls_layout.addWidget(self.page_label)
        
        # Preview widget
        self.preview_widget = QPrintPreviewWidget(self.printer, self)
        self.preview_widget.paintRequested.connect(self.render_for_preview)
        
        # Navigation buttons
        nav_layout = QHBoxLayout()
        
        self.first_page_btn = QPushButton("First Page")
        self.prev_page_btn = QPushButton("Previous Page")
        self.next_page_btn = QPushButton("Next Page")
        self.last_page_btn = QPushButton("Last Page")
        
        nav_layout.addWidget(self.first_page_btn)
        nav_layout.addWidget(self.prev_page_btn)
        nav_layout.addWidget(self.next_page_btn)
        nav_layout.addWidget(self.last_page_btn)
        
        # Add to preview layout
        preview_layout.addLayout(controls_layout)
        preview_layout.addWidget(self.preview_widget, 1)
        preview_layout.addLayout(nav_layout)
        preview_group.setLayout(preview_layout)
        
        # Button box
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Help
        )
        
        # Add Print button (custom, not in standard button box roles)
        self.print_btn = button_box.addButton("Print", QDialogButtonBox.AcceptRole)
        self.print_btn.setDefault(True)
        
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        # Add to main layout
        main_layout.addWidget(options_widget, 1)
        main_layout.addWidget(preview_group, 2)
        
        # Add button box at the bottom
        main_layout.addWidget(button_box, 0, Qt.AlignBottom)
        
        # Connect signals
        self.setup_connections()
        
        # Initial preview update
        self.update_preview()
    
    def setup_connections(self):
        """Setup signal connections."""
        # Navigation buttons
        self.first_page_btn.clicked.connect(self.goto_first_page)
        self.prev_page_btn.clicked.connect(self.goto_prev_page)
        self.next_page_btn.clicked.connect(self.goto_next_page)
        self.last_page_btn.clicked.connect(self.goto_last_page)
        
        # Zoom combo
        self.zoom_combo.currentTextChanged.connect(self.update_zoom)
        
        # Page range changes
        self.range_all.toggled.connect(self.update_preview)
        self.range_current.toggled.connect(self.update_preview)
        self.range_selection.toggled.connect(self.update_preview)
        self.range_custom.toggled.connect(self.update_preview)
        self.range_edit.textChanged.connect(self.update_preview)
        
        # Other options
        self.print_to_file.toggled.connect(self.update_preview)
    
    def update_preview(self):
        """Update the print preview."""
        self.preview_widget.updatePreview()
    
    def render_for_preview(self, printer):
        """Render the document for the print preview."""
        # Configure the printer based on dialog settings
        self.configure_printer(printer)
        
        # Create a painter to render to the printer
        painter = QPainter()
        painter.begin(printer)
        
        # Get the page layout
        page_rect = printer.pageRect(QPrinter.DevicePixel)
        
        # Create a temporary document for rendering
        doc = QTextDocument()
        doc.setDocumentMargin(0)
        
        # For demonstration, we'll just show the page numbers
        # In a real implementation, you would render the actual document content
        html = """
        <html>
        <head>
            <style>
                body { 
                    font-family: Arial; 
                    margin: 0; 
                    padding: 20px;
                    height: 100%;
                    box-sizing: border-box;
                }
                .page {
                    border: 1px solid #ccc;
                    height: 100%;
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                    align-items: center;
                }
                .page-number {
                    font-size: 24px;
                    font-weight: bold;
                    margin-bottom: 20px;
                }
                .content {
                    text-align: center;
                    color: #666;
                }
            </style>
        </head>
        <body>
            <div class="page">
                <div class="page-number">Page 1</div>
                <div class="content">
                    <p>This is a preview of how the document will be printed.</p>
                    <p>In the actual implementation, this would show the real document content.</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        doc.setHtml(html)
        doc.setPageSize(page_rect.size())
        
        # Render the document
        doc.drawContents(painter)
        
        # End painting
        painter.end()
    
    def configure_printer(self, printer):
        """Configure the printer based on dialog settings."""
        # Set number of copies
        printer.setCopyCount(self.copies_spin.value())
        
        # Set collation
        printer.setCollateCopies(self.collate_check.isChecked())
        
        # Set page range
        if self.range_all.isChecked():
            printer.setPrintRange(QPrinter.AllPages)
        elif self.range_current.isChecked():
            printer.setPrintRange(QPrinter.CurrentPage)
        elif self.range_selection.isChecked():
            printer.setPrintRange(QPrinter.Selection)
        else:  # Custom range
            printer.setPrintRange(QPrinter.PageRange)
            # Parse the page range (simplified)
            # In a real implementation, you would parse the range string properly
            pages = self.range_edit.text().strip()
            if pages:
                try:
                    # Simple implementation - just get the first page number
                    first_page = int(pages.split('-')[0].split(',')[0].strip())
                    printer.setFromTo(first_page, first_page)
                except (ValueError, IndexError):
                    pass
    
    def show_printer_properties(self):
        """Show printer properties dialog."""
        dialog = QPrintDialog(self.printer, self)
        if dialog.exec() == QDialog.Accepted:
            self.update_preview()
    
    def goto_first_page(self):
        """Go to the first page in the preview."""
        self.preview_widget.setCurrentPage(0)
    
    def goto_prev_page(self):
        """Go to the previous page in the preview."""
        current = self.preview_widget.currentPage()
        if current > 0:
            self.preview_widget.setCurrentPage(current - 1)
    
    def goto_next_page(self):
        """Go to the next page in the preview."""
        current = self.preview_widget.currentPage()
        self.preview_widget.setCurrentPage(current + 1)
    
    def goto_last_page(self):
        """Go to the last page in the preview."""
        # In a real implementation, you would know the total number of pages
        # For now, just go to a high page number
        self.preview_widget.setCurrentPage(999)
    
    def update_zoom(self, zoom_text):
        """Update the zoom level of the preview."""
        if zoom_text.endswith('%'):
            try:
                zoom = float(zoom_text[:-1]) / 100.0
                self.preview_widget.setZoomFactor(zoom)
            except ValueError:
                pass
        elif zoom_text == "Page Width":
            self.preview_widget.fitToWidth()
        elif zoom_text == "Text Width":
            self.preview_widget.fitInView()
        elif zoom_text == "Whole Page":
            self.preview_widget.fitToPage()
    
    def accept(self):
        """Handle dialog acceptance."""
        # In a real implementation, you would print the document here
        print("Printing...")  # Placeholder
        super().accept()
    
    @staticmethod
    def get_print_dialog(document, parent=None):
        """Static method to show the print dialog."""
        dialog = PrintDialog(document, parent)
        return dialog.exec() == QDialog.Accepted
