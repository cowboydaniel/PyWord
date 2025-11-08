from PySide6.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog
from PySide6.QtCore import QMarginsF, QSizeF, Qt
from PySide6.QtGui import QPainter, QPageSize, QPageLayout, QTextDocument
from typing import Optional
import os
from .page_setup import PageOrientation

class PrintManager:
    """Handles printing and print preview functionality."""
    
    def __init__(self, parent=None):
        self.parent = parent
        self.printer = QPrinter(QPrinter.HighResolution)
        self.printer.setFullPage(True)
        
        # Set default page size to A4
        self.page_layout = QPageLayout(
            QPageSize(QPageSize.PageSizeId.A4),
            QPageLayout.Orientation.Portrait,
            QMarginsF(72, 72, 72, 72),  # 1 inch margins
            QPageLayout.Unit.Point
        )
        self.printer.setPageLayout(self.page_layout)
    
    def print_document(self, document: QTextDocument) -> bool:
        """Print the document using system print dialog."""
        if not document:
            return False
            
        dialog = QPrintDialog(self.printer, self.parent)
        if dialog.exec() != QPrintDialog.Accepted:
            return False
            
        return self._print_document(document)
    
    def print_preview(self, document: QTextDocument):
        """Show print preview dialog."""
        if not document:
            return
            
        preview = QPrintPreviewDialog(self.printer, self.parent)
        preview.paintRequested.connect(lambda printer: self._print_document(document, printer))
        preview.exec()
    
    def print_to_pdf(self, document: QTextDocument, file_path: str) -> bool:
        """Export document to PDF."""
        if not document or not file_path:
            return False
            
        # Ensure the file has a .pdf extension
        if not file_path.lower().endswith('.pdf'):
            file_path += '.pdf'
            
        # Use a temporary printer for PDF export
        pdf_printer = QPrinter(QPrinter.HighResolution)
        pdf_printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        pdf_printer.setOutputFileName(file_path)
        pdf_printer.setPageLayout(self.page_layout)
        
        return self._print_document(document, pdf_printer)
    
    def _print_document(self, document: QTextDocument, printer: Optional[QPrinter] = None) -> bool:
        """Internal method to handle actual printing."""
        target_printer = printer or self.printer
        
        try:
            # Save current page size and margins
            page_size = target_printer.pageLayout().pageSize()
            margins = target_printer.pageLayout().margins(QPageLayout.Unit.Point)
            
            # Set document page size
            document.setPageSize(QSizeF(page_size.sizePoints()))
            
            # Set document margins
            document.setDocumentMargin(min(margins.left(), margins.top()))
            
            # Print the document
            document.print_(target_printer)
            return True
            
        except Exception as e:
            print(f"Error printing document: {e}")
            return False
    
    def set_page_setup(self, page_setup: 'PageSetup'):
        """Update printer settings from PageSetup."""
        # Create page size
        page_size = QPageSize(
            QSizeF(page_setup.page_width, page_setup.page_height),
            QPageSize.Unit.Point
        )
        
        # Create margins
        margins = QMarginsF(
            page_setup.margins.left,
            page_setup.margins.top,
            page_setup.margins.right,
            page_setup.margins.bottom
        )
        
        # Set page orientation
        orientation = QPageLayout.Orientation.Portrait
        if page_setup.orientation == PageOrientation.LANDSCAPE:
            orientation = QPageLayout.Orientation.Landscape

        # Update page layout
        self.page_layout = QPageLayout(
            page_size,
            orientation,
            margins,
            QPageLayout.Unit.Point
        )
        
        # Apply to printer
        self.printer.setPageLayout(self.page_layout)
