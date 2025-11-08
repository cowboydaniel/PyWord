"""
Document Management UI Components for PyWord.

This module provides UI components for managing documents, including:
- Document tabs
- Document switching
- File operations (New, Open, Save, Save As)
- Recent documents
"""

from typing import Optional, Callable, Dict, Any, List
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, QToolBar,
    QMenu, QFileDialog, QMessageBox, QLabel, QSizePolicy
)
from PySide6.QtCore import Qt, Signal, Slot, QSize
from PySide6.QtGui import QIcon, QKeySequence, QPixmap, QAction

from ..core.document import Document, DocumentType, DocumentManager
from .toolbars.main_toolbar import MainToolBar

class DocumentTabWidget(QTabWidget):
    """
    A tab widget for managing multiple document editors.
    """
    # Signals
    current_editor_changed = Signal(object)  # Emits the current editor widget
    tab_close_requested = Signal(int)        # Emits tab index to close
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTabsClosable(True)
        self.setMovable(True)
        self.setDocumentMode(True)
        self.setElideMode(Qt.TextElideMode.ElideRight)

        # Hide tab bar for Microsoft Word style (single document interface)
        self.tabBar().setVisible(False)

        # Connect signals
        self.tabCloseRequested.connect(self.on_tab_close_requested)
        self.currentChanged.connect(self.on_current_changed)
    
    def add_editor(self, editor: QWidget, title: str, tooltip: str = "") -> int:
        """
        Add a new editor tab.
        
        Args:
            editor: The editor widget to add
            title: Tab title
            tooltip: Optional tooltip text
            
        Returns:
            int: The index of the new tab
        """
        index = self.addTab(editor, title)
        self.setTabToolTip(index, tooltip)
        self.setCurrentIndex(index)
        return index
    
    def current_editor(self) -> Optional[QWidget]:
        """Get the current editor widget."""
        return self.currentWidget()
    
    def set_current_editor(self, editor: QWidget):
        """Set the current editor widget."""
        index = self.indexOf(editor)
        if index >= 0:
            self.setCurrentIndex(index)
    
    def editor_at(self, index: int) -> Optional[QWidget]:
        """Get the editor widget at the specified index."""
        return self.widget(index)
    
    def find_editor(self, document_path: str) -> Optional[QWidget]:
        """
        Find an editor by document path.
        
        Args:
            document_path: Path to the document
            
        Returns:
            The editor widget if found, None otherwise
        """
        if not document_path:
            return None
            
        for i in range(self.count()):
            editor = self.widget(i)
            if hasattr(editor, 'document_path') and editor.document_path == document_path:
                return editor
        return None
    
    @Slot(int)
    def on_tab_close_requested(self, index: int):
        """Handle tab close request."""
        self.tab_close_requested.emit(index)
    
    @Slot(int)
    def on_current_changed(self, index: int):
        """Handle tab change."""
        self.current_editor_changed.emit(self.widget(index))


class DocumentManagerUI(QWidget):
    """
    Main document management UI component.
    
    This class provides the UI for managing documents, including:
    - Document tabs
    - File operations (New, Open, Save, Save As)
    - Recent documents
    - Document switching
    """
    # Signals
    document_activated = Signal(object)      # Emits the activated document
    document_closed = Signal(object)         # Emits the closed document
    document_saved = Signal(object, str)     # Emits (document, file_path)
    
    def __init__(self, document_manager: DocumentManager, parent=None):
        super().__init__(parent)
        self.document_manager = document_manager
        self.current_document = None
        self.recent_docs_menu = None
        
        self.setup_ui()
        self.setup_connections()
    
    def setup_ui(self):
        """Initialize the UI components."""
        # Main layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Create toolbar (hidden for Word-like interface - using ribbon instead)
        self.toolbar = self.create_toolbar()
        self.toolbar.setVisible(False)
        layout.addWidget(self.toolbar)

        # Create tab widget for documents
        self.tab_widget = DocumentTabWidget()
        layout.addWidget(self.tab_widget)

        # Status bar (hidden - using main window status bar instead)
        self.status_bar = QLabel()
        self.status_bar.setFrameStyle(QLabel.StyledPanel | QLabel.Sunken)
        self.status_bar.setStyleSheet("padding: 2px;")
        self.status_bar.setVisible(False)
        layout.addWidget(self.status_bar)

        # Update UI state
        self.update_ui_state()
    
    def create_toolbar(self) -> QToolBar:
        """Create the document management toolbar."""
        toolbar = QToolBar("Document")
        toolbar.setIconSize(QSize(16, 16))
        
        # New Document
        self.new_action = QAction("New", self)
        self.new_action.setIcon(QIcon.fromTheme("document-new"))
        self.new_action.setShortcut(QKeySequence.New)
        self.new_action.setToolTip("Create a new document")
        toolbar.addAction(self.new_action)
        
        # Open Document
        self.open_action = QAction("Open...", self)
        self.open_action.setIcon(QIcon.fromTheme("document-open"))
        self.open_action.setShortcut(QKeySequence.Open)
        self.open_action.setToolTip("Open an existing document")
        toolbar.addAction(self.open_action)
        
        # Save
        self.save_action = QAction("Save", self)
        self.save_action.setIcon(QIcon.fromTheme("document-save"))
        self.save_action.setShortcut(QKeySequence.Save)
        self.save_action.setToolTip("Save the current document")
        toolbar.addAction(self.save_action)
        
        # Save As
        self.save_as_action = QAction("Save As...", self)
        self.save_as_action.setIcon(QIcon.fromTheme("document-save-as"))
        self.save_as_action.setShortcut(QKeySequence.SaveAs)
        self.save_as_action.setToolTip("Save the current document with a new name")
        toolbar.addAction(self.save_as_action)
        
        toolbar.addSeparator()

        # Recent Documents menu
        self.recent_menu = QMenu("Recent Documents", self)
        self.recent_docs_action = self.recent_menu.menuAction()
        self.recent_docs_action.setIcon(QIcon.fromTheme("document-open-recent"))
        self.recent_docs_action.setText("Recent")
        toolbar.addAction(self.recent_docs_action)
        
        # Update recent documents
        self.update_recent_documents_menu()
        
        return toolbar
    
    def setup_connections(self):
        """Set up signal connections."""
        # Toolbar actions
        self.new_action.triggered.connect(self.new_document)
        self.open_action.triggered.connect(self.open_document_dialog)
        self.save_action.triggered.connect(self.save_document)
        self.save_as_action.triggered.connect(self.save_document_as)
        
        # Tab widget signals
        self.tab_widget.current_editor_changed.connect(self.on_current_editor_changed)
        self.tab_widget.tab_close_requested.connect(self.close_tab)
        
        # Document manager signals
        self.document_manager.add_listener(self.on_document_event)
    
    def update_ui_state(self):
        """Update the UI state based on the current document."""
        has_document = self.current_document is not None
        self.save_action.setEnabled(has_document)
        self.save_as_action.setEnabled(has_document)
    
    def update_recent_documents_menu(self):
        """Update the recent documents menu."""
        self.recent_menu.clear()
        
        recent_docs = self.document_manager.get_recent_documents()
        if not recent_docs:
            action = self.recent_menu.addAction("No recent documents")
            action.setEnabled(False)
            return
        
        for doc_info in recent_docs:
            path = doc_info.get('path', '')
            title = doc_info.get('title', Path(path).name)
            
            action = self.recent_menu.addAction(title)
            action.setData(path)
            action.triggered.connect(
                lambda checked, p=path: self.open_document(p)
            )
        
        self.recent_menu.addSeparator()
        clear_action = self.recent_menu.addAction("Clear Recent Documents")
        clear_action.triggered.connect(self.document_manager.clear_recent_documents)
    
    @Slot()
    def new_document(self):
        """Create a new document."""
        # TODO: Show template selection dialog
        doc = self.document_manager.create_document()
        self.open_document_editor(doc)
    
    @Slot()
    def open_document_dialog(self):
        """Show the open document dialog."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Document",
            "",
            "All Supported Formats (*.txt *.rtf *.html *.htm *.docx *.odt);;"
            "Text Files (*.txt);;"
            "Rich Text Files (*.rtf);;"
            "HTML Files (*.html *.htm);;"
            "Word Documents (*.docx);;"
            "OpenDocument Text (*.odt);;"
            "All Files (*.*)"
        )
        
        if file_path:
            self.open_document(file_path)
    
    def open_document(self, file_path: str):
        """
        Open a document from the specified file path.
        
        Args:
            file_path: Path to the document to open
        """
        # Check if document is already open
        editor = self.tab_widget.find_editor(file_path)
        if editor:
            self.tab_widget.setCurrentWidget(editor)
            return
        
        # Open the document
        doc = self.document_manager.open_document(file_path)
        if doc:
            self.open_document_editor(doc)
    
    def open_document_editor(self, document: Document):
        """
        Open an editor for the specified document.
        
        Args:
            document: The document to open an editor for
        """
        # Check if editor already exists for this document
        if document.file_path:
            editor = self.tab_widget.find_editor(document.file_path)
            if editor:
                self.tab_widget.setCurrentWidget(editor)
                return
        
        # Create a new editor
        from ..core.editor import TextEditor  # Lazy import to avoid circular imports
        editor = TextEditor()
        
        # Set document content
        if document.content:
            editor.setPlainText(document.content)
        
        # Add to tab widget
        title = document.title or "Untitled"
        tooltip = document.file_path or ""
        self.tab_widget.add_editor(editor, title, tooltip)
        
        # Store document reference in editor
        editor.pyword_document = document
        if document.file_path:
            editor.document_path = document.file_path

        # Connect editor signals
        editor.document().contentsChanged.connect(
            lambda: self.on_document_modified(editor, document)
        )
        
        # Update current document
        self.current_document = document
        self.update_ui_state()
    
    @Slot()
    def save_document(self):
        """Save the current document."""
        if not self.current_document:
            return
        
        if not self.current_document.file_path:
            self.save_document_as()
        else:
            if self.document_manager.save_document():
                self.status_bar.setText(f"Document saved: {self.current_document.file_path}")
    
    @Slot()
    def save_document_as(self):
        """Save the current document with a new name."""
        if not self.current_document:
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Document As",
            "",
            "Text Files (*.txt);;Rich Text Files (*.rtf);;HTML Files (*.html);;Word Documents (*.docx);;OpenDocument Text (*.odt)"
        )
        
        if file_path:
            if self.document_manager.save_document_as(file_path):
                self.status_bar.setText(f"Document saved as: {file_path}")
    
    @Slot(int)
    def close_tab(self, index: int):
        """
        Close the tab at the specified index.
        
        Args:
            index: Index of the tab to close
        """
        editor = self.tab_widget.editor_at(index)
        if not editor:
            return
        
        # Check for unsaved changes
        document = getattr(editor, 'pyword_document', None)
        if document and document.modified:
            reply = QMessageBox.question(
                self,
                "Unsaved Changes",
                f"Save changes to {document.title or 'Untitled'}?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                QMessageBox.Save
            )
            
            if reply == QMessageBox.Save:
                self.save_document()
                return  # Don't close if save was cancelled
            elif reply == QMessageBox.Cancel:
                return
        
        # Remove the tab
        self.tab_widget.removeTab(index)
        
        # Clean up
        if hasattr(editor, 'deleteLater'):
            editor.deleteLater()
        
        # Update current document
        if document:
            self.document_manager.close_document(document)
            self.document_closed.emit(document)
    
    @Slot(object)
    def on_current_editor_changed(self, editor: QWidget):
        """Handle editor tab change."""
        if not editor:
            self.current_document = None
            self.update_ui_state()
            return

        document = getattr(editor, 'pyword_document', None)
        if document != self.current_document:
            self.current_document = document
            self.document_activated.emit(document)
            self.update_ui_state()
    
    def on_document_modified(self, editor: QWidget, document: Document):
        """Handle document modification."""
        if not document or not hasattr(editor, 'pyword_document'):
            return
        
        # Update document content
        document.content = editor.toPlainText()
        document.modified = True
        
        # Update tab title to show modified state
        index = self.tab_widget.indexOf(editor)
        if index >= 0:
            title = self.tab_widget.tabText(index)
            if not title.endswith('*'):
                self.tab_widget.setTabText(index, f"{title}*")
    
    def on_document_event(self, event_type: str, **kwargs):
        """Handle document manager events."""
        if event_type == 'document_created':
            doc = kwargs.get('document')
            if doc:
                self.open_document_editor(doc)
        
        elif event_type == 'document_opened':
            doc = kwargs.get('document')
            if doc:
                self.open_document_editor(doc)
        
        elif event_type == 'document_saved':
            doc = kwargs.get('document')
            if doc:
                # Update tab title to remove modified marker
                editor = self.tab_widget.find_editor(doc.file_path)
                if editor:
                    index = self.tab_widget.indexOf(editor)
                    if index >= 0:
                        title = self.tab_widget.tabText(index).rstrip('*')
                        self.tab_widget.setTabText(index, title)
        
        elif event_type == 'recent_documents_cleared':
            self.update_recent_documents_menu()
        
        # Update UI based on document state
        self.update_ui_state()


# Example usage
if __name__ == "__main__":
    import sys
    from PySide6.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    
    # Create document manager
    doc_manager = DocumentManager()
    
    # Create document manager UI
    doc_ui = DocumentManagerUI(doc_manager)
    doc_ui.setWindowTitle("Document Manager")
    doc_ui.resize(800, 600)
    doc_ui.show()
    
    sys.exit(app.exec())
