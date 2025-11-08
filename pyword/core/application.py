from PySide6.QtWidgets import (QMainWindow, QVBoxLayout, QWidget, QStatusBar,
                             QMenuBar, QMenu, QToolBar, QFileDialog, QMessageBox,
                             QSplitter, QDockWidget, QListWidget, QListWidgetItem,
                             QTextEdit, QDockWidget, QTextBrowser, QDialog)
from PySide6.QtGui import QFont
from PySide6.QtCore import QSettings, Qt, QSize, QTimer
from PySide6.QtGui import QAction, QKeySequence, QIcon, QTextDocument, QTextCursor
import os
import sys

# Import features
from ..features.styles import DocumentStyles
from ..features.tables import TableManager
from ..features.navigation import FindReplaceDialog, DocumentMap, GoToDialog
from ..features.headers_footers import HeaderFooterManager, HeaderFooterType
from ..features.columns import ColumnManager, ColumnSettings
from ..features.lists import ListManager, ListType
from ..features.page_numbers import PageNumberManager, PageNumberSettings
from ..features.shapes import ShapeManager
from ..features.split_view import SplitViewManager
from ..features.performance import (LargeDocumentOptimizer, MemoryManager,
                                   BackgroundSaver, AutoRecovery, PerformanceMonitor)
from ..features.accessibility import AccessibilityManager, AccessibilityLevel

# Import UI components
from ..ui.ribbon import RibbonBar
from ..ui.theme_manager import ThemeManager, Theme

class WordProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.documents = []
        self.current_document = None
        self.current_file = None
        self.setWindowTitle("Document1 - PyWord")
        self.setGeometry(100, 100, 1200, 800)

        # Initialize settings
        self.settings = QSettings("PyWord", "Editor")

        # Initialize and apply theme (Microsoft Word style)
        self.theme_manager = ThemeManager(self)
        self.theme_manager.apply_theme()
        
        # Initialize features
        self.styles = DocumentStyles()
        self.table_manager = TableManager(self)

        # Initialize header/footer manager (will be set when document is loaded)
        self.header_footer_manager = None

        # Initialize additional feature managers
        self.column_manager = None  # Will be initialized when text_edit is ready
        self.list_manager = None
        self.page_number_manager = None
        self.shape_manager = None
        self.split_view_manager = None

        # Initialize Phase 6 features - Performance
        self.document_optimizer = LargeDocumentOptimizer()
        self.memory_manager = MemoryManager()
        self.background_saver = BackgroundSaver()
        self.auto_recovery = AutoRecovery()
        self.performance_monitor = PerformanceMonitor()

        # Initialize Phase 6 features - Accessibility
        self.accessibility_manager = AccessibilityManager(self)

        # Setup UI
        self.setup_ui()

        # Initialize document map
        self.init_document_map()

        # Initialize all feature managers now that text_edit is available
        self.init_feature_managers()

        # Start Phase 6 features
        self.init_phase6_features()

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
        # Create main container widget
        container = QWidget()
        container.setObjectName("centralWidget")
        self.setCentralWidget(container)

        # Main layout
        main_layout = QVBoxLayout(container)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Create ribbon interface (Microsoft Word style)
        self.ribbon = RibbonBar(self)
        main_layout.addWidget(self.ribbon)

        # Create and add ribbon tabs
        home_tab = self.ribbon.create_home_tab()
        insert_tab = self.ribbon.create_insert_tab()
        design_tab = self.ribbon.create_design_tab()
        layout_tab = self.ribbon.create_layout_tab()
        view_tab = self.ribbon.create_view_tab()

        self.ribbon.add_tab(home_tab)
        self.ribbon.add_tab(insert_tab)
        self.ribbon.add_tab(design_tab)
        self.ribbon.add_tab(layout_tab)
        self.ribbon.add_tab(view_tab)

        # Create menu bar
        self.create_menus()

        # Create document area with splitter
        self.splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(self.splitter)

        # Create a basic text editor for demonstration
        self.text_edit = QTextEdit()
        self.text_edit.setStyleSheet("QTextEdit { background-color: #F3F2F1; }")
        self.splitter.addWidget(self.text_edit)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
    
    def load_settings(self):
        """Load application settings."""
        if self.settings.contains("geometry"):
            self.restoreGeometry(self.settings.value("geometry"))
        if self.settings.contains("windowState"):
            self.restoreState(self.settings.value("windowState"))

    def init_phase6_features(self):
        """Initialize Phase 6 features (Performance and Accessibility)."""
        # Start auto-recovery
        self.auto_recovery.start()

        # Start background saver
        self.background_saver.start()

        # Connect performance monitor warnings
        self.performance_monitor.performance_warning.connect(
            lambda msg: self.status_bar.showMessage(f"Performance: {msg}", 5000)
        )

        # Connect auto-recovery signals
        self.auto_recovery.recovery_available.connect(self.handle_recovery_files)

        # Setup accessibility features
        self.accessibility_manager.screen_reader.announcement_requested.connect(
            lambda msg, priority: self.status_bar.showMessage(msg, 3000)
        )

        # Enable basic accessibility by default
        self.accessibility_manager.set_accessibility_level(AccessibilityLevel.BASIC)

    def handle_recovery_files(self, recovery_files):
        """Handle available recovery files."""
        if not recovery_files:
            return

        # Show recovery dialog
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QListWidget
        dialog = QDialog(self)
        dialog.setWindowTitle("Recover Documents")
        dialog.setMinimumSize(400, 300)

        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel("The following documents can be recovered:"))

        list_widget = QListWidget()
        for recovery in recovery_files:
            title = recovery.get('title', 'Untitled')
            timestamp = recovery.get('timestamp', '')
            list_widget.addItem(f"{title} - {timestamp}")

        layout.addWidget(list_widget)

        button_layout = QVBoxLayout()
        recover_btn = QPushButton("Recover Selected")
        skip_btn = QPushButton("Skip")

        def recover_selected():
            current_row = list_widget.currentRow()
            if current_row >= 0:
                recovery_file = recovery_files[current_row]['file']
                recovered = self.auto_recovery.recover_document(recovery_file)
                if recovered:
                    self.status_bar.showMessage(f"Recovered: {recovered['title']}", 5000)
            dialog.accept()

        recover_btn.clicked.connect(recover_selected)
        skip_btn.clicked.connect(dialog.reject)

        button_layout.addWidget(recover_btn)
        button_layout.addWidget(skip_btn)
        layout.addLayout(button_layout)

        dialog.exec()

    def closeEvent(self, event):
        """Handle application close event."""
        # Save window state
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())

        # Stop Phase 6 features
        self.auto_recovery.stop()
        self.background_saver.stop()
        self.accessibility_manager.tts.stop()

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
    
    # Ribbon interface replaces traditional toolbars
    # Old toolbar code removed - using Microsoft Word-style ribbon instead
    
    def text_bold(self):
        """Toggle bold formatting for the selected text."""
        if hasattr(self, 'text_edit') and self.text_edit:
            fmt = self.text_edit.currentCharFormat()
            weight = QFont.Weight.Normal if fmt.fontWeight() > QFont.Weight.Normal else QFont.Weight.Bold
            fmt.setFontWeight(weight)
            self.text_edit.mergeCurrentCharFormat(fmt)
    
    def init_feature_managers(self):
        """Initialize all feature managers after text_edit is created."""
        if not hasattr(self, 'text_edit') or not self.text_edit:
            return

        # Column manager
        self.column_manager = ColumnManager(self.text_edit.document())

        # List manager
        self.list_manager = ListManager(self.text_edit)

        # Page number manager
        self.page_number_manager = PageNumberManager(self.text_edit.document())

        # Shape manager
        self.shape_manager = ShapeManager(self.text_edit.document())

        # Split view manager
        self.split_view_manager = SplitViewManager(self)

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

        menu.addSeparator()

        # Split view actions
        split_horizontal_action = QAction("Split &Horizontal", self)
        split_horizontal_action.triggered.connect(self.split_view_horizontal)
        menu.addAction(split_horizontal_action)

        split_vertical_action = QAction("Split &Vertical", self)
        split_vertical_action.triggered.connect(self.split_view_vertical)
        menu.addAction(split_vertical_action)

        remove_split_action = QAction("&Remove Split", self)
        remove_split_action.triggered.connect(self.remove_split_view)
        menu.addAction(remove_split_action)
    
    def setup_insert_actions(self, menu):
        """Setup insert-related actions."""
        table_action = QAction("&Table...", self)
        table_action.triggered.connect(self.insert_table)
        menu.addAction(table_action)

        menu.addSeparator()

        # Page numbers
        page_numbers_action = QAction("&Page Numbers...", self)
        page_numbers_action.triggered.connect(self.insert_page_numbers)
        menu.addAction(page_numbers_action)

        menu.addSeparator()

        # Shapes submenu
        shapes_menu = menu.addMenu("&Shapes")

        rectangle_action = QAction("Rectangle", self)
        rectangle_action.triggered.connect(lambda: self.insert_shape("rectangle"))
        shapes_menu.addAction(rectangle_action)

        circle_action = QAction("Circle", self)
        circle_action.triggered.connect(lambda: self.insert_shape("circle"))
        shapes_menu.addAction(circle_action)

        line_action = QAction("Line", self)
        line_action.triggered.connect(lambda: self.insert_shape("line"))
        shapes_menu.addAction(line_action)

        arrow_action = QAction("Arrow", self)
        arrow_action.triggered.connect(lambda: self.insert_shape("arrow"))
        shapes_menu.addAction(arrow_action)
    
    def setup_format_actions(self, menu):
        """Setup format-related actions."""
        # Bullets and numbering
        bullets_action = QAction("&Bullets", self)
        bullets_action.triggered.connect(lambda: self.apply_list_format(ListType.BULLET))
        menu.addAction(bullets_action)

        numbering_action = QAction("&Numbering", self)
        numbering_action.triggered.connect(lambda: self.apply_list_format(ListType.NUMBERED))
        menu.addAction(numbering_action)

        menu.addSeparator()

        # Columns
        columns_action = QAction("C&olumns...", self)
        columns_action.triggered.connect(self.format_columns)
        menu.addAction(columns_action)

        menu.addSeparator()

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
    
    def apply_list_format(self, list_type):
        """Apply list formatting (bullets or numbering)."""
        if self.list_manager:
            if list_type == ListType.BULLET:
                self.list_manager.create_bullet_list()
            elif list_type == ListType.NUMBERED:
                self.list_manager.create_numbered_list()

    def format_columns(self):
        """Show columns dialog and apply column formatting."""
        if self.column_manager:
            from ..features.columns import ColumnDialog, ColumnSettings
            dialog = ColumnDialog(self)
            if dialog.exec():
                settings = dialog.get_settings()
                self.column_manager.apply_columns(settings)

    def insert_page_numbers(self):
        """Insert page numbers."""
        if self.page_number_manager:
            from ..features.page_numbers import PageNumberDialog
            dialog = PageNumberDialog(self)
            if dialog.exec():
                settings = dialog.get_settings()
                self.page_number_manager.insert_page_numbers(settings)

    def insert_shape(self, shape_type):
        """Insert a shape."""
        if self.shape_manager:
            # Default size for shapes
            width, height = 100, 100
            if shape_type == "line":
                self.shape_manager.insert_line(width)
            elif shape_type == "rectangle":
                self.shape_manager.insert_rectangle(width, height)
            elif shape_type == "circle":
                self.shape_manager.insert_ellipse(width, height)
            elif shape_type == "arrow":
                self.shape_manager.insert_arrow(width)

    def split_view_horizontal(self):
        """Split view horizontally."""
        if self.split_view_manager:
            self.split_view_manager.split_horizontal()

    def split_view_vertical(self):
        """Split view vertically."""
        if self.split_view_manager:
            self.split_view_manager.split_vertical()

    def remove_split_view(self):
        """Remove split view."""
        if self.split_view_manager:
            self.split_view_manager.remove_split()

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
