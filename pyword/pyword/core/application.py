from PySide6.QtWidgets import (QMainWindow, QVBoxLayout, QWidget, QStatusBar, 
                             QMenuBar, QMenu, QToolBar, QFileDialog, QMessageBox,
                             QSplitter, QDockWidget, QListWidget, QListWidgetItem,
                             QTextEdit, QDockWidget, QTextBrowser)
from PySide6.QtGui import QFont
from PySide6.QtCore import QSettings, Qt, QSize, QTimer
from PySide6.QtGui import QAction, QKeySequence, QIcon, QTextDocument
import os
import sys

# Import features
from ..features.styles import DocumentStyles
from ..features.tables import TableManager
from ..features.navigation import FindReplaceDialog, DocumentMap, GoToDialog
from ..features.headers_footers import HeaderFooterManager, HeaderFooterType

# Import UI components
from ..ui.toolbars.header_footer_toolbar import HeaderFooterToolBar

class WordProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.documents = []
        self.current_document = None
        self.current_file = None
        self.setWindowTitle("PyWord")
        self.setGeometry(100, 100, 1200, 800)
        
        # Initialize settings
        self.settings = QSettings("PyWord", "Editor")
        
        # Initialize features
        self.styles = DocumentStyles()
        self.table_manager = TableManager(self)
        
        # Initialize header/footer manager (will be set when document is loaded)
        self.header_footer_manager = None
        
        # Setup UI
        self.setup_ui()
        
        # Initialize document map
        self.init_document_map()
        
        # Initialize header/footer toolbar
        self.header_footer_toolbar = HeaderFooterToolBar(self)
        self.header_footer_toolbar.setObjectName("headerFooterToolbar")
        self.addToolBar(Qt.TopToolBarArea, self.header_footer_toolbar)
        self.header_footer_toolbar.hide()  # Hide by default, show when document is loaded
        
        # Load settings
        self.load_settings()
    
    def new_file(self):
        """Create a new document."""
        # In a real implementation, this would create a new document
        # For now, we'll just update the window title
        self.current_file = None
        self.setWindowTitle("Untitled - PyWord")
        
    def open_file(self):
        """Open an existing document."""
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Open File", "", "Text Files (*.txt);;All Files (*)")
            
        if file_name:
            try:
                with open(file_name, 'r') as file:
                    content = file.read()
                    # In a real implementation, you would set this content to the editor
                    self.current_file = file_name
                    self.setWindowTitle(f"{os.path.basename(file_name)} - PyWord")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not open file: {str(e)}")
    
    def save_file(self):
        """Save the current document."""
        if not self.current_file:
            return self.save_file_as()
            
        try:
            # In a real implementation, get content from the editor
            content = ""  # This would be the actual content from the editor
            with open(self.current_file, 'w') as file:
                file.write(content)
            return True
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not save file: {str(e)}")
            return False
    
    def save_file_as(self):
        """Save the current document with a new name."""
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save File", "", "Text Files (*.txt);;All Files (*)")
            
        if file_name:
            self.current_file = file_name
            self.setWindowTitle(f"{os.path.basename(file_name)} - PyWord")
            return self.save_file()
        return False
    
    def setup_ui(self):
        """Initialize the main UI components."""
        # Central widget and splitter
        self.splitter = QSplitter(Qt.Horizontal)
        
        # Create document area
        self.central_widget = QWidget()
        self.setCentralWidget(self.splitter)
        
        # Main layout
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setContentsMargins(0, 0, 0, 0)
        
        # Create a basic text editor for demonstration
        self.text_edit = QTextEdit()
        self.splitter.addWidget(self.text_edit)
        
        # Set initial window title
        self.setWindowTitle("Untitled - PyWord")
        self.layout.setSpacing(0)
        
        # Create menu bar
        self.create_menus()
        
        # Create toolbars
        self.create_toolbars()
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        
        # Add widgets to splitter
        self.splitter.addWidget(self.central_widget)
    
    def load_settings(self):
        """Load application settings."""
        if self.settings.contains("geometry"):
            self.restoreGeometry(self.settings.value("geometry"))
        if self.settings.contains("windowState"):
            self.restoreState(self.settings.value("windowState"))
    
    def closeEvent(self, event):
        """Handle application close event."""
        # Save window state
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())
        
        # Handle unsaved documents
        if self.maybe_save():
            event.accept()
        else:
            event.ignore()
    
    def create_menus(self):
        """Create the application menu bar."""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("&File")
        
        # Edit menu
        edit_menu = menubar.addMenu("&Edit")
        
        # View menu
        view_menu = menubar.addMenu("&View")
        
        # Insert menu
        insert_menu = menubar.addMenu("&Insert")
        
        # Format menu
        format_menu = menubar.addMenu("F&ormat")
        
        # Table menu
        table_menu = menubar.addMenu("&Table")
        
        # Add actions to menus
        self.setup_file_actions(file_menu)
        self.setup_edit_actions(edit_menu)
        self.setup_view_actions(view_menu)
        self.setup_insert_actions(insert_menu)
        self.setup_format_actions(format_menu)
        self.setup_table_actions(table_menu)
    
    def create_toolbars(self):
        """Create application toolbars."""
        # Formatting toolbar
        self.format_toolbar = self.addToolBar("Format")
        self.format_toolbar.setObjectName("formatToolbar")
        self.format_toolbar.setIconSize(QSize(16, 16))
        
        # Add formatting actions
        bold_action = QAction(QIcon.fromTheme("format-text-bold"), "Bold", self)
        bold_action.triggered.connect(self.text_bold)
        self.format_toolbar.addAction(bold_action)
        
        # Add more formatting actions as needed...
    
    def text_bold(self):
        """Toggle bold formatting for the selected text."""
        if hasattr(self, 'text_edit') and self.text_edit:
            fmt = self.text_edit.currentCharFormat()
            weight = QFont.Weight.Normal if fmt.fontWeight() > QFont.Weight.Normal else QFont.Weight.Bold
            fmt.setFontWeight(weight)
            self.text_edit.mergeCurrentCharFormat(fmt)
    
    def init_document_map(self):
        """Initialize the document map dock widget."""
        self.dock = QDockWidget("Document Map", self)
        self.dock.setObjectName("documentMapDock")
        self.dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.document_map = DocumentMap(self.text_edit)
        self.dock.setWidget(self.document_map)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock)
        self.dock.hide()  # Hidden by default
    
    def toggle_document_map(self):
        """Toggle document map visibility."""
        self.dock.setVisible(not self.dock.isVisible())
    
    def show_find_dialog(self):
        """Show the find/replace dialog."""
        dialog = FindReplaceDialog(self)
        dialog.show()
    
    def show_go_to_dialog(self):
        """Show the go to dialog."""
        # Calculate total pages (simplified)
        page_count = max(1, self.text_edit.document().blockCount() // 50)
        dialog = GoToDialog(page_count, self)
        if dialog.exec() == QDialog.Accepted:
            page = dialog.get_page_number()
            if page:
                # Simplified page navigation
                cursor = self.text_edit.textCursor()
                cursor.movePosition(QTextCursor.Start)
                cursor.movePosition(QTextCursor.Down, QTextCursor.MoveAnchor, (page - 1) * 50)
                self.text_edit.setTextCursor(cursor)
    
    def insert_table(self):
        """Insert a table at the current cursor position."""
        # This would typically show a dialog to select rows/columns
        self.table_manager.insert_table(3, 3)
    
    def setup_file_actions(self, menu):
        """Setup file-related actions."""
        new_action = QAction("&New", self)
        new_action.setShortcut(QKeySequence.New)
        new_action.triggered.connect(self.new_file)
        menu.addAction(new_action)
        
        # Add other file actions...
    
    def setup_edit_actions(self, menu):
        """Setup edit-related actions."""
        find_action = QAction("&Find...", self)
        find_action.setShortcut(QKeySequence.Find)
        find_action.triggered.connect(self.show_find_dialog)
        menu.addAction(find_action)
        
        go_to_action = QAction("&Go To...", self)
        go_to_action.setShortcut("Ctrl+G")
        go_to_action.triggered.connect(self.show_go_to_dialog)
        menu.addAction(go_to_action)
    
    def setup_view_actions(self, menu):
        """Setup view-related actions."""
        doc_map_action = QAction("&Document Map", self)
        doc_map_action.setCheckable(True)
        doc_map_action.setChecked(False)
        doc_map_action.triggered.connect(self.toggle_document_map)
        menu.addAction(doc_map_action)
    
    def setup_insert_actions(self, menu):
        """Setup insert-related actions."""
        table_action = QAction("&Table...", self)
        table_action.triggered.connect(self.insert_table)
        menu.addAction(table_action)
    
    def setup_format_actions(self, menu):
        """Setup format-related actions."""
        # Add style actions
        for style in self.styles.styles.keys():
            action = QAction(style, self)
            action.triggered.connect(lambda checked, s=style: self.apply_style(s))
            menu.addAction(action)
        
        # Add theme submenu
        theme_menu = menu.addMenu("Themes")
        for theme in self.styles.themes.keys():
            action = QAction(theme, self, checkable=True)
            action.setChecked(theme == self.styles.current_theme)
            action.triggered.connect(lambda checked, t=theme: self.set_theme(t))
            theme_menu.addAction(action)
    
    def setup_table_actions(self, menu):
        """Setup table-related actions."""
        insert_row_action = QAction("Insert &Row", self)
        insert_row_action.triggered.connect(self.table_manager.insert_row)
        menu.addAction(insert_row_action)
        
        insert_col_action = QAction("Insert &Column", self)
        insert_col_action.triggered.connect(self.table_manager.insert_column)
        menu.addAction(insert_col_action)
        
        menu.addSeparator()
        
        delete_row_action = QAction("Delete Ro&w", self)
        delete_row_action.triggered.connect(self.table_manager.delete_row)
        menu.addAction(delete_row_action)
        
        delete_col_action = QAction("Delete Col&umn", self)
        delete_col_action.triggered.connect(self.table_manager.delete_column)
        menu.addAction(delete_col_action)
    
    def apply_style(self, style_name):
        """Apply the selected style to the current selection."""
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            cursor.mergeCharFormat(self.styles.get_style(style_name))
    
    def set_theme(self, theme_name):
        """Apply the selected theme."""
        if self.styles.set_theme(theme_name):
            colors = self.styles.get_theme_colors()
            self.setStyleSheet(f"""
                QMainWindow, QWidget {{
                    background-color: {colors['background']};
                    color: {colors['text']};
                }}
                QTextEdit {{
                    background-color: {colors['background']};
                    color: {colors['text']};
                    selection-background-color: {colors['highlight']};
                }}
            """)
    
    def maybe_save(self):
        """Check if the document has unsaved changes."""
        if self.text_edit.document().isModified():
            ret = QMessageBox.warning(
                self, "Document Modified",
                "The document has been modified.\nDo you want to save your changes?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
            )
            
            if ret == QMessageBox.Save:
                return self.save_file()
            elif ret == QMessageBox.Cancel:
                return False
        return True
    
    def new_document(self):
        """Create a new document."""
        # TODO: Implement new document creation
        pass
    
    def open_document(self, file_path=None):
        """Open an existing document."""
        # TODO: Implement document opening
        pass
    
    def save_document(self):
        """Save the current document."""
        # TODO: Implement document saving
        pass
    
    def save_document_as(self):
        """Save the current document with a new name."""
        # TODO: Implement save as functionality
        pass
