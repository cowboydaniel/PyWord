"""Print Preview Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QToolBar,
                             QPushButton, QLabel, QSpinBox, QComboBox)
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QIcon, QPageLayout
from PySide6.QtPrintSupport import QPrintPreviewWidget, QPrinter
from .base_dialog import BaseDialog


class PrintPreviewDialog(BaseDialog):
    """Dialog for previewing documents before printing."""

    def __init__(self, document=None, parent=None):
        # Initialize attributes before calling super().__init__()
        # because BaseDialog.__init__() calls setup_ui()
        self.document = document
        self.printer = QPrinter(QPrinter.HighResolution)
        super().__init__("Print Preview", parent)

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Toolbar
        toolbar = QToolBar()

        # Print button
        print_action = QAction(QIcon.fromTheme("document-print"), "Print", self)
        print_action.triggered.connect(self.print_document)
        toolbar.addAction(print_action)

        toolbar.addSeparator()

        # Zoom controls
        zoom_in_action = QAction(QIcon.fromTheme("zoom-in"), "Zoom In", self)
        zoom_in_action.triggered.connect(self.zoom_in)
        toolbar.addAction(zoom_in_action)

        zoom_out_action = QAction(QIcon.fromTheme("zoom-out"), "Zoom Out", self)
        zoom_out_action.triggered.connect(self.zoom_out)
        toolbar.addAction(zoom_out_action)

        toolbar.addSeparator()

        # Page navigation
        toolbar.addWidget(QLabel("Page:"))
        self.page_spinbox = QSpinBox()
        self.page_spinbox.setMinimum(1)
        self.page_spinbox.setValue(1)
        toolbar.addWidget(self.page_spinbox)

        layout.addWidget(toolbar)

        # Preview widget
        self.preview = QPrintPreviewWidget(self.printer)
        if self.document:
            self.preview.paintRequested.connect(self.document.print_)
        layout.addWidget(self.preview)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.accept)
        button_layout.addWidget(self.close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.resize(800, 600)

    def print_document(self):
        """Print the document."""
        # TODO: Show print dialog and print
        pass

    def zoom_in(self):
        """Zoom in the preview."""
        self.preview.zoomIn()

    def zoom_out(self):
        """Zoom out the preview."""
        self.preview.zoomOut()
