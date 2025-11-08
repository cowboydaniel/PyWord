from datetime import datetime
from PySide6.QtWidgets import (QMainWindow, QVBoxLayout, QWidget, QSplitter, QTabWidget,
                             QStatusBar, QMenuBar, QMenu, QToolBar, QFileDialog, QMessageBox,
                             QDockWidget, QLabel, QSizePolicy, QApplication, QDialog, QVBoxLayout,
                             QHBoxLayout, QPushButton, QTextEdit, QLineEdit, QListWidget, QTreeView,
                             QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QSpinBox,
                             QDoubleSpinBox, QCheckBox, QGroupBox, QFormLayout, QScrollArea,
                             QFrame, QToolButton, QStyle, QColorDialog, QFontDialog, QInputDialog,
                             QSplitterHandle, QPrintDialog, QPrintPreviewDialog)
from PySide6.QtPrintSupport import QPrinter, QPageSetupDialog
from PySide6.QtCore import Qt, QSize, QSettings, QTimer, QUrl, QMimeData, Signal
from PySide6.QtGui import (QAction, QIcon, QFont, QTextCursor, QTextCharFormat, QTextListFormat,
                         QTextBlockFormat, QTextTableFormat, QTextFrameFormat, QTextLength,
                         QTextImageFormat, QPixmap, QImage, QPainter, QPalette, QColor,
                         QDesktopServices, QKeySequence, QTextDocument, QTextFrame, QTextTable,
                         QTextDocumentFragment, QTextBlock, QTextList, QTextFormat, QTextFrameLayoutData,
                         QShortcut)

from ...core.document import Document, DocumentManager
from ...core.editor import TextEditor
from ...core.page_setup import PageSetup, PageOrientation, PageMargins
from ...core.print_manager import PrintManager
from .document_manager_ui import DocumentManagerUI
from .toolbars import MainToolBar, FormatToolBar, TableToolBar, ReviewToolBar, ViewToolBar
from .panels import NavigationPanel, StylesPanel, DocumentMapPanel, CommentsPanel
from .ribbon import RibbonBar
from .theme_manager import ThemeManager, Theme
from .dialogs import (NewDocumentDialog, PageSetupDialog, PrintPreviewDialog, InsertTableDialog,
                     InsertImageDialog, InsertLinkDialog, FindReplaceDialog, WordCountDialog,
                     GoToDialog, OptionsDialog, AboutDialog, StyleDialog, TablePropertiesDialog,
                     BulletsAndNumberingDialog, BorderAndShadingDialog, ColumnsDialog,
                     TabsDialog, ParagraphDialog, FontDialog, SymbolDialog, HyperlinkDialog)

class MainWindow(QMainWindow):
    # Signals
    zoom_changed = Signal(int)  # Emits zoom percentage
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Document1 - PyWord")
        self.setGeometry(100, 100, 1200, 800)
        self.current_zoom = 100  # Store current zoom level

        # Initialize core components
        self.document_manager = DocumentManager()
        self.settings = QSettings("PyWord", "Editor")
        self.print_manager = PrintManager(self)

        # Initialize and apply theme (Microsoft Word style)
        self.theme_manager = ThemeManager(self)
        self.theme_manager.apply_theme()

        # Setup UI
        self.setup_ui()
        
        # Connect print actions
        self.setup_print_connections()
        
        # Load settings
        self.load_settings()
        
        # Create document manager UI
        self.document_ui = DocumentManagerUI(self.document_manager, self)
        
        # Connect document UI signals
        self.document_ui.document_activated.connect(self.on_document_activated)
        self.document_ui.document_closed.connect(self.on_document_closed)
        self.document_ui.document_saved.connect(self.on_document_saved)
        
        # Connect text changes to update word count
        if self.current_editor():
            self.current_editor().textChanged.connect(self.update_word_count)
        
        # Setup zoom shortcuts
        self.setup_zoom_shortcuts()
        
        # Create a new document by default if none exists
        if not self.document_manager.documents:
            self.document_ui.new_document()
    
    def setup_print_connections(self):
        """Connect print-related signals and slots."""
        # Connect print actions
        self.actionPrint.triggered.connect(self.print_document)
        self.actionPrint_Preview.triggered.connect(self.print_preview)
        self.actionPage_Setup.triggered.connect(self.page_setup)
        self.actionPrint.triggered.connect(self.print_document)
        
    def print_document(self):
        """Print the current document."""
        if not self.current_editor():
            return
            
        # Update printer settings from document
        if self.current_document:
            self.print_manager.set_page_setup(self.current_document.page_setup)
            
        # Get the document content as QTextDocument
        doc = self.current_editor().document()
        
        # Show print dialog and print
        self.print_manager.print_document(doc)
    
    def print_preview(self):
        """Show print preview for the current document."""
        if not self.current_editor():
            return
            
        # Update printer settings from document
        if self.current_document:
            self.print_manager.set_page_setup(self.current_document.page_setup)
            
        # Get the document content as QTextDocument
        doc = self.current_editor().document()
        
        # Show print preview
        self.print_manager.print_preview(doc)
    
    def page_setup(self):
        """Show page setup dialog and update document settings."""
        if not self.current_document:
            return
            
        # Create and show page setup dialog
        dialog = PageSetupDialog(self.current_document.page_setup, self)
        if dialog.exec() == QDialog.Accepted:
            # Update document's page setup
            self.current_document.page_setup = dialog.get_page_setup()
            self.current_document.modified = True
            
            # Update printer settings
            self.print_manager.set_page_setup(self.current_document.page_setup)
            
            # TODO: Update document view to reflect new page setup
            
    def update_word_count(self):
        """Update the word count in the status bar."""
        if not self.current_editor():
            return
            
        text = self.current_editor().toPlainText()
        word_count = len(text.split())
        char_count = len(text)
        self.word_count_label.setText(f"Words: {word_count} | Chars: {char_count}")
    
    def update_zoom_display(self, zoom_level: int):
        """Update the zoom level display in the status bar."""
        self.current_zoom = zoom_level
        self.zoom_label.setText(f"{zoom_level}%")
    
    def setup_zoom_shortcuts(self):
        """Setup keyboard shortcuts for zooming."""
        # Zoom in: Ctrl++
        zoom_in = QShortcut(QKeySequence("Ctrl++"), self)
        zoom_in.activated.connect(self.zoom_in)
        
        # Zoom out: Ctrl+-
        zoom_out = QShortcut(QKeySequence("Ctrl+-"), self)
        zoom_out.activated.connect(self.zoom_out)
        
        # Reset zoom: Ctrl+0
        zoom_reset = QShortcut(QKeySequence("Ctrl+0"), self)
        zoom_reset.activated.connect(self.zoom_reset)
    
    def zoom_in(self):
        """Zoom in the current editor."""
        if self.current_editor():
            self.current_editor().zoom_in()
    
    def zoom_out(self):
        """Zoom out the current editor."""
        if self.current_editor():
            self.current_editor().zoom_out()
    
    def zoom_reset(self):
        """Reset zoom to 100% in the current editor."""
        if self.current_editor():
            self.current_editor().zoom_reset()
    
    def setup_ui(self):
        """Initialize the main window UI."""
        # Create central widget and layout
        self.central_widget = QWidget()
        self.central_widget.setObjectName("centralWidget")
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        # Create toolbars
        self.setup_toolbars()
        
        # Create main splitter for editor and side panels
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_layout.addWidget(self.main_splitter)
        
        # Add document manager UI to the main layout
        self.main_splitter.addWidget(self.document_ui)
        
        # Create left panel (navigation, styles, etc.)
        self.setup_left_panel()
        
        # Create editor area
        self.setup_editor_area()
        
        # Create right panel (document map, comments, etc.)
        self.setup_right_panel()
        
        # Create status bar
        self.setup_status_bar()
        
        # Create menus
        self.setup_menus()
        
        # Update UI
        self.update_ui()
    
    def setup_toolbars(self):
        """Create and setup the ribbon interface."""
        # Create ribbon interface (Microsoft Word style)
        self.ribbon = RibbonBar(self)
        self.main_layout.addWidget(self.ribbon)

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

        # Connect ribbon actions to methods
        self.connect_ribbon_actions()

    def connect_ribbon_actions(self):
        """Connect ribbon button actions to their respective methods."""
        # This method will connect ribbon buttons to the application methods
        # For now, we'll let the ribbon display without connections
        # Full connections can be added incrementally
        pass

    def setup_left_panel(self):
        """Setup the left panel with navigation and styles."""
        self.left_panel = QDockWidget("Navigation", self)
        self.left_panel.setObjectName("leftPanel")
        self.left_panel.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable)
        
        left_panel_widget = QWidget()
        left_panel_layout = QVBoxLayout(left_panel_widget)
        left_panel_layout.setContentsMargins(0, 0, 0, 0)
        left_panel_layout.setSpacing(0)
        
        # Tab widget for different panels
        self.left_tab_widget = QTabWidget()
        self.left_tab_widget.setTabPosition(QTabWidget.West)
        
        # Navigation panel
        self.navigation_panel = NavigationPanel(self)
        self.left_tab_widget.addTab(self.navigation_panel, "Navigation")
        
        # Styles panel
        self.styles_panel = StylesPanel(self)
        self.left_tab_widget.addTab(self.styles_panel, "Styles")
        
        left_panel_layout.addWidget(self.left_tab_widget)
        self.left_panel.setWidget(left_panel_widget)
        
        self.addDockWidget(Qt.LeftDockWidgetArea, self.left_panel)
    
    def setup_editor_area(self):
        """Setup the main editor area."""
        self.editor_area = QWidget()
        editor_layout = QVBoxLayout(self.editor_area)
        editor_layout.setContentsMargins(0, 0, 0, 0)
        editor_layout.setSpacing(0)
        
        # Tab widget for multiple documents
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.setMovable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_tab)
        self.tab_widget.currentChanged.connect(self.tab_changed)
        
        editor_layout.addWidget(self.tab_widget)
        self.main_splitter.addWidget(self.editor_area)
    
    def setup_right_panel(self):
        """Setup the right panel with document map and comments."""
        self.right_panel = QDockWidget("Document Map", self)
        self.right_panel.setObjectName("rightPanel")
        self.right_panel.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable)
        
        right_panel_widget = QWidget()
        right_panel_layout = QVBoxLayout(right_panel_widget)
        right_panel_layout.setContentsMargins(0, 0, 0, 0)
        right_panel_layout.setSpacing(0)
        
        # Tab widget for different panels
        self.right_tab_widget = QTabWidget()
        self.right_tab_widget.setTabPosition(QTabWidget.East)
        
        # Document map panel
        self.document_map_panel = DocumentMapPanel(self)
        self.right_tab_widget.addTab(self.document_map_panel, "Document Map")
        
        # Comments panel
        self.comments_panel = CommentsPanel(self)
        self.right_tab_widget.addTab(self.comments_panel, "Comments")
        
        right_panel_layout.addWidget(self.right_tab_widget)
        self.right_panel.setWidget(right_panel_widget)
        
        self.addDockWidget(Qt.RightDockWidgetArea, self.right_panel)
    
    def setup_status_bar(self):
        """Setup the status bar."""
        # Create status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        
        # Add status bar widgets
        self.word_count_label = QLabel("Words: 0")
        self.page_info_label = QLabel("Page: 1")
        self.zoom_label = QLabel("100%")
        
        self.status_bar.addPermanentWidget(self.word_count_label, 1)
        self.status_bar.addPermanentWidget(self.page_info_label, 1)
        self.status_bar.addPermanentWidget(self.zoom_label, 0)
        
        # Connect zoom signal
        self.zoom_changed.connect(self.update_zoom_display)
        
        # Page info
        self.page_info = QLabel("Page 1 of 1")
        self.status_bar.addPermanentWidget(self.page_info)
        
        # Word count
        self.word_count = QLabel("Words: 0")
        self.status_bar.addPermanentWidget(self.word_count)
        
        # Zoom level
        self.zoom_level = QLabel("100%")
        self.status_bar.addPermanentWidget(self.zoom_level)
    
    def setup_menus(self):
        """Setup all menus."""
        # File menu
        file_menu = self.menuBar().addMenu("&File")
        
        new_action = QAction("&New", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_document)
        file_menu.addAction(new_action)
        
        open_action = QAction("&Open...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_document)
        file_menu.addAction(open_action)
        
        save_action = QAction("&Save", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_document)
        file_menu.addAction(save_action)
        
        save_as_action = QAction("Save &As...", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_document_as)
        file_menu.addAction(save_as_action)
        
        file_menu.addSeparator()
        
        page_setup_action = QAction("Page Set&up...", self)
        page_setup_action.triggered.connect(self.page_setup)
        file_menu.addAction(page_setup_action)
        
        print_preview_action = QAction("Print Pre&view", self)
        print_preview_action.triggered.connect(self.print_preview)
        file_menu.addAction(print_preview_action)
        
        print_action = QAction("&Print...", self)
        print_action.setShortcut("Ctrl+P")
        print_action.triggered.connect(self.print_document)
        file_menu.addAction(print_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("E&xit", self)
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Edit menu
        edit_menu = self.menuBar().addMenu("&Edit")
        
        # Undo action
        self.actionUndo = QAction("&Undo", self)
        self.actionUndo.setShortcut("Ctrl+Z")
        self.actionUndo.triggered.connect(self.undo)
        self.actionUndo.setEnabled(False)
        edit_menu.addAction(self.actionUndo)
        
        # Redo action
        self.actionRedo = QAction("&Redo", self)
        self.actionRedo.setShortcut("Ctrl+Y")
        self.actionRedo.triggered.connect(self.redo)
        self.actionRedo.setEnabled(False)
        edit_menu.addAction(self.actionRedo)
        
        edit_menu.addSeparator()
        
        # Connect document change signals
        if self.current_editor():
            self.current_editor().undoAvailable.connect(self.actionUndo.setEnabled)
            self.current_editor().redoAvailable.connect(self.actionRedo.setEnabled)
        
        cut_action = QAction("Cu&t", self)
        cut_action.setShortcut("Ctrl+X")
        cut_action.triggered.connect(self.cut)
        edit_menu.addAction(cut_action)
        
        copy_action = QAction("&Copy", self)
        copy_action.setShortcut("Ctrl+C")
        copy_action.triggered.connect(self.copy)
        edit_menu.addAction(copy_action)
        
        paste_action = QAction("&Paste", self)
        paste_action.setShortcut("Ctrl+V")
        paste_action.triggered.connect(self.paste)
        edit_menu.addAction(paste_action)
        
        edit_menu.addSeparator()
        
        find_action = QAction("&Find...", self)
        find_action.setShortcut("Ctrl+F")
        find_action.triggered.connect(self.find)
        edit_menu.addAction(find_action)
        
        replace_action = QAction("&Replace...", self)
        replace_action.setShortcut("Ctrl+H")
        replace_action.triggered.connect(self.replace)
        edit_menu.addAction(replace_action)
        
        goto_action = QAction("&Go To...", self)
        goto_action.setShortcut("Ctrl+G")
        goto_action.triggered.connect(self.go_to)
        edit_menu.addAction(goto_action)
        
        # View menu
        view_menu = self.menuBar().addMenu("&View")
        
        zoom_in_action = QAction("Zoom &In", self)
        zoom_in_action.setShortcut("Ctrl++")
        zoom_in_action.triggered.connect(self.zoom_in)
        view_menu.addAction(zoom_in_action)
        
        zoom_out_action = QAction("Zoom &Out", self)
        zoom_out_action.setShortcut("Ctrl+-")
        zoom_out_action.triggered.connect(self.zoom_out)
        view_menu.addAction(zoom_out_action)
        
        zoom_reset_action = QAction("Reset &Zoom", self)
        zoom_reset_action.setShortcut("Ctrl+0")
        zoom_reset_action.triggered.connect(self.zoom_reset)
        view_menu.addAction(zoom_reset_action)
        
        view_menu.addSeparator()
        
        # Toolbars submenu
        toolbars_menu = view_menu.addMenu("&Toolbars")
        
        standard_toolbar_action = QAction("&Standard", self)
        standard_toolbar_action.setCheckable(True)
        standard_toolbar_action.setChecked(True)
        standard_toolbar_action.triggered.connect(self.toggle_standard_toolbar)
        toolbars_menu.addAction(standard_toolbar_action)
        
        formatting_toolbar_action = QAction("&Formatting", self)
        formatting_toolbar_action.setCheckable(True)
        formatting_toolbar_action.setChecked(True)
        formatting_toolbar_action.triggered.connect(self.toggle_formatting_toolbar)
        toolbars_menu.addAction(formatting_toolbar_action)
        
        # Panels submenu
        panels_menu = view_menu.addMenu("&Panels")
        
        navigation_panel_action = QAction("&Navigation", self)
        navigation_panel_action.setCheckable(True)
        navigation_panel_action.setChecked(True)
        navigation_panel_action.triggered.connect(self.toggle_navigation_panel)
        panels_menu.addAction(navigation_panel_action)
        
        styles_panel_action = QAction("&Styles", self)
        styles_panel_action.setCheckable(True)
        styles_panel_action.setChecked(True)
        styles_panel_action.triggered.connect(self.toggle_styles_panel)
        panels_menu.addAction(styles_panel_action)
        
        document_map_action = QAction("Document &Map", self)
        document_map_action.setCheckable(True)
        document_map_action.setChecked(True)
        document_map_action.triggered.connect(self.toggle_document_map)
        panels_menu.addAction(document_map_action)
        
        comments_panel_action = QAction("Co&mments", self)
        comments_panel_action.setCheckable(True)
        comments_panel_action.setChecked(True)
        comments_panel_action.triggered.connect(self.toggle_comments_panel)
        panels_menu.addAction(comments_panel_action)
        
        # Insert menu
        insert_menu = self.menuBar().addMenu("&Insert")
        
        page_break_action = QAction("Page &Break", self)
        page_break_action.setShortcut("Ctrl+Enter")
        page_break_action.triggered.connect(self.insert_page_break)
        insert_menu.addAction(page_break_action)
        
        table_action = QAction("&Table...", self)
        table_action.triggered.connect(self.insert_table)
        insert_menu.addAction(table_action)
        
        image_action = QAction("&Picture...", self)
        image_action.triggered.connect(self.insert_image)
        insert_menu.addAction(image_action)
        
        hyperlink_action = QAction("&Hyperlink...", self)
        hyperlink_action.setShortcut("Ctrl+K")
        hyperlink_action.triggered.connect(self.insert_hyperlink)
        insert_menu.addAction(hyperlink_action)
        
        symbol_action = QAction("&Symbol...", self)
        symbol_action.triggered.connect(self.insert_symbol)
        insert_menu.addAction(symbol_action)
        
        # Format menu
        format_menu = self.menuBar().addMenu("F&ormat")
        
        font_action = QAction("&Font...", self)
        font_action.triggered.connect(self.format_font)
        format_menu.addAction(font_action)
        
        paragraph_action = QAction("&Paragraph...", self)
        paragraph_action.triggered.connect(self.format_paragraph)
        format_menu.addAction(paragraph_action)
        
        bullets_action = QAction("&Bullets and Numbering...", self)
        bullets_action.triggered.connect(self.format_bullets)
        format_menu.addAction(bullets_action)
        
        border_action = QAction("&Borders and Shading...", self)
        border_action.triggered.connect(self.format_borders)
        format_menu.addAction(border_action)
        
        columns_action = QAction("Colu&mns...", self)
        columns_action.triggered.connect(self.format_columns)
        format_menu.addAction(columns_action)
        
        tabs_action = QAction("&Tabs...", self)
        tabs_action.triggered.connect(self.format_tabs)
        format_menu.addAction(tabs_action)
        
        # Table menu
        self.table_menu = self.menuBar().addMenu("&Table")
        self.setup_table_menu()
        
        # Tools menu
        tools_menu = self.menuBar().addMenu("&Tools")
        
        word_count_action = QAction("Word &Count", self)
        word_count_action.triggered.connect(self.show_word_count)
        tools_menu.addAction(word_count_action)
        
        tools_menu.addSeparator()
        
        options_action = QAction("&Options...", self)
        options_action.triggered.connect(self.show_options)
        tools_menu.addAction(options_action)
        
        # Help menu
        help_menu = self.menuBar().addMenu("&Help")
        
        about_action = QAction("&About PyWord", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def setup_table_menu(self):
        """Setup the table menu with table-related actions."""
        self.table_menu.clear()
        
        insert_table_action = QAction("Insert &Table...", self)
        insert_table_action.triggered.connect(self.insert_table)
        self.table_menu.addAction(insert_table_action)
        
        delete_table_action = QAction("&Delete Table", self)
        delete_table_action.triggered.connect(self.delete_table)
        self.table_menu.addAction(delete_table_action)
        
        self.table_menu.addSeparator()
        
        insert_row_above_action = QAction("Insert Rows &Above", self)
        insert_row_above_action.triggered.connect(lambda: self.insert_table_row(True))
        self.table_menu.addAction(insert_row_above_action)
        
        insert_row_below_action = QAction("Insert Rows &Below", self)
        insert_row_below_action.triggered.connect(lambda: self.insert_table_row(False))
        self.table_menu.addAction(insert_row_below_action)
        
        insert_column_left_action = QAction("Insert &Columns to the Left", self)
        insert_column_left_action.triggered.connect(lambda: self.insert_table_column(True))
        self.table_menu.addAction(insert_column_left_action)
        
        insert_column_right_action = QAction("Insert Columns to the &Right", self)
        insert_column_right_action.triggered.connect(lambda: self.insert_table_column(False))
        self.table_menu.addAction(insert_column_right_action)
        
        self.table_menu.addSeparator()
        
        delete_row_action = QAction("Delete &Rows", self)
        delete_row_action.triggered.connect(self.delete_table_row)
        self.table_menu.addAction(delete_row_action)
        
        delete_column_action = QAction("Delete &Columns", self)
        delete_column_action.triggered.connect(self.delete_table_column)
        self.table_menu.addAction(delete_column_action)
        
        self.table_menu.addSeparator()
        
        merge_cells_action = QAction("&Merge Cells", self)
        merge_cells_action.triggered.connect(self.merge_table_cells)
        self.table_menu.addAction(merge_cells_action)
        
        split_cells_action = QAction("S&plit Cells...", self)
        split_cells_action.triggered.connect(self.split_table_cells)
        self.table_menu.addAction(split_cells_action)
        
        self.table_menu.addSeparator()
        
        table_properties_action = QAction("Table P&roperties...", self)
    def update_ui(self):
        """Update the UI based on the current state."""
        has_document = self.document_ui.current_document is not None
        
        # Update menu and toolbar states
        for action in self.findChildren(QAction):
            if action not in [a for a in self.menuBar().actions()]:
                action.setEnabled(has_document)
        
        # Update status bar
        self.update_status_bar()

    def update_status_bar(self):
        """Update the status bar with current document information."""
        if not hasattr(self, 'status_bar'):
            return
            
        doc = self.document_ui.current_document
        if not doc:
            self.status_bar.clear()
            return
            
        # Get document statistics
        stats = doc.get_metadata()
        
        # Format status text
        status_text = []
        
        # Document name
        if doc.file_path:
            status_text.append(f"{Path(doc.file_path).name}")
        else:
            status_text.append("Untitled")
            
        # Modified indicator
        if doc.modified:
            status_text[-1] += " *"
            
        # Page/word count
        status_text.append(f"Words: {stats.get('word_count', 0)}")
        status_text.append(f"Chars: {stats.get('char_count', 0)}")
        
        # Update status bar
        self.status_bar.setText(" | ".join(status_text))

    # Document management
    def new_document(self):
        """Create a new document."""
        self.document_ui.new_document()

    def open_document(self):
        """Open an existing document."""
        self.document_ui.open_document()

    def save_document(self):
        """Save the current document."""
        self.document_ui.save_document()

    def save_document_as(self):
        """Save the current document with a new name."""
        self.document_ui.save_document_as()

    def close_document(self, index=None):
        """Close the current or specified document."""
        if index is None:
            index = self.document_ui.tab_widget.currentIndex()
        if index >= 0:
            self.document_ui.close_tab(index)

    def closeEvent(self, event):
        """Handle window close event."""
        # Check for unsaved changes
        if self.document_manager.has_unsaved_changes:
            reply = QMessageBox.question(
                self,
                "Unsaved Changes",
                "You have unsaved changes. Do you want to save them before exiting?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                QMessageBox.Save
            )
            
            if reply == QMessageBox.Save:
                if not self.document_ui.save_document():
                    event.ignore()
                    return
            elif reply == QMessageBox.Cancel:
                event.ignore()
                return
        
        # Save window state and geometry
        self.settings.setValue("window/geometry", self.saveGeometry())
        self.settings.setValue("window/state", self.saveState())
        
        # Save recent documents
        recent_docs = [{
            'path': doc.get('path', ''),
            'title': doc.get('title', '')
        } for doc in self.document_manager.get_recent_documents()]
        self.settings.setValue("recent_documents", recent_docs)
        
        event.accept()

    def get_current_editor(self):
        """Get the current active editor."""
        return self.document_ui.tab_widget.current_editor()

    def current_editor(self):
        """Get the current active editor (alias for get_current_editor)."""
        return self.get_current_editor()

    def tab_changed(self, index):
        """Handle tab change event."""
        if index >= 0:
            editor = self.tab_widget.widget(index)
            if hasattr(editor, 'document'):
                self.current_document = editor.document
            else:
                self.current_document = None
        else:
            self.current_document = None
        
        self.update_ui()
    
    def close_tab(self, index):
        """Close the tab at the given index."""
        self.close_document(index)
    
    # Document operations
    def document_modified(self):
        """Handle document modification."""
        editor = self.get_current_editor()
        if editor and hasattr(editor, 'document'):
            editor.document.content = editor.toPlainText()
            editor.document.modified = True
            
            # Update tab text to show modification indicator
            index = self.tab_widget.indexOf(editor)
            if index >= 0:
                title = editor.document.title
                if not title.endswith('*'):
                    self.tab_widget.setTabText(index, f"{title}*")
            
            self.update_ui()
    
    # Formatting methods
    def format_font(self):
        """Open font dialog to format text."""
        editor = self.get_current_editor()
        if editor:
            font, ok = QFontDialog.getFont(editor.currentFont(), self, "Select Font")
            if ok:
                editor.setCurrentFont(font)
    
    def format_paragraph(self):
        """Open paragraph formatting dialog."""
        editor = self.get_current_editor()
        if editor:
            dialog = ParagraphDialog(editor, self)
            if dialog.exec() == QDialog.Accepted:
                # Apply paragraph formatting
                pass
    
    def format_bullets(self):
        """Open bullets and numbering dialog."""
        dialog = BulletsAndNumberingDialog(self)
        if dialog.exec() == QDialog.Accepted:
            # Apply bullet/numbering formatting
            pass
    
    def format_borders(self):
        """Open borders and shading dialog."""
        dialog = BorderAndShadingDialog(self)
        if dialog.exec() == QDialog.Accepted:
            # Apply border/shading formatting
            pass
    
    def format_columns(self):
        """Open columns dialog."""
        dialog = ColumnsDialog(self)
        if dialog.exec() == QDialog.Accepted:
            # Apply column formatting
            pass
    
    def format_tabs(self):
        """Open tabs dialog."""
        dialog = TabsDialog(self)
        if dialog.exec() == QDialog.Accepted:
            # Apply tab settings
            pass
    
    # Insert methods
    def insert_page_break(self):
        """Insert a page break at the current cursor position."""
        editor = self.get_current_editor()
        if editor:
            cursor = editor.textCursor()
            cursor.insertHtml("<hr style='page-break-before:always; margin:0;' />")
    
    def insert_table(self):
        """Insert a table at the current cursor position."""
        editor = self.get_current_editor()
        if editor:
            dialog = InsertTableDialog(self)
            if dialog.exec() == QDialog.Accepted:
                rows = dialog.get_rows()
                cols = dialog.get_columns()
                editor.insert_table(rows, cols)
    
    def insert_image(self):
        """Insert an image at the current cursor position."""
        editor = self.get_current_editor()
        if editor:
            dialog = InsertImageDialog(self)
            if dialog.exec() == QDialog.Accepted:
                image_path = dialog.get_image_path()
                if image_path:
                    editor.insert_image(image_path)
    
    def insert_hyperlink(self):
        """Insert a hyperlink at the current cursor position."""
        editor = self.get_current_editor()
        if editor:
            dialog = HyperlinkDialog(self)
            if dialog.exec() == QDialog.Accepted:
                url = dialog.get_url()
                text = dialog.get_text()
                if url and text:
                    editor.insertHtml(f'<a href="{url}">{text}</a>')
    
    def insert_symbol(self):
        """Insert a symbol at the current cursor position."""
        dialog = SymbolDialog(self)
        if dialog.exec() == QDialog.Accepted:
            symbol = dialog.get_selected_symbol()
            if symbol:
                editor = self.get_current_editor()
                if editor:
                    editor.insertPlainText(symbol)
    
    # Table operations
    def delete_table(self):
        """Delete the current table."""
        editor = self.get_current_editor()
        if editor:
            cursor = editor.textCursor()
            if cursor.currentTable():
                cursor.currentTable().removeRows(0, cursor.currentTable().rows())
    
    def insert_table_row(self, above):
        """Insert a row above or below the current row."""
        editor = self.get_current_editor()
        if editor:
            cursor = editor.textCursor()
            table = cursor.currentTable()
            if table:
                if above:
                    table.insertRows(cursor.currentRow(), 1)
                else:
                    table.insertRows(cursor.currentRow() + 1, 1)
    
    def insert_table_column(self, left):
        """Insert a column to the left or right of the current column."""
        editor = self.get_current_editor()
        if editor:
            cursor = editor.textCursor()
            table = cursor.currentTable()
            if table:
                if left:
                    table.insertColumns(cursor.columnNumber(), 1)
                else:
                    table.insertColumns(cursor.columnNumber() + 1, 1)
    
    def delete_table_row(self):
        """Delete the current row."""
        editor = self.get_current_editor()
        if editor:
            cursor = editor.textCursor()
            table = cursor.currentTable()
            if table:
                table.removeRows(cursor.currentRow(), 1)
    
    def delete_table_column(self):
        """Delete the current column."""
        editor = self.get_current_editor()
        if editor:
            cursor = editor.textCursor()
            table = cursor.currentTable()
            if table:
                table.removeColumns(cursor.columnNumber(), 1)
    
    def merge_table_cells(self):
        """Merge selected table cells."""
        editor = self.get_current_editor()
        if editor:
            cursor = editor.textCursor()
            if cursor.hasSelection() and cursor.currentTable():
                cursor.currentTable().mergeCells(cursor)
    
    def split_table_cells(self):
        """Split the current table cell."""
        editor = self.get_current_editor()
        if editor:
            cursor = editor.textCursor()
            if cursor.currentTable():
                dialog = QInputDialog(self)
                rows, ok1 = dialog.getInt(self, "Split Cells", "Number of rows:", 1, 1, 100, 1)
                cols, ok2 = dialog.getInt(self, "Split Cells", "Number of columns:", 1, 1, 100, 1)
                
                if ok1 and ok2 and (rows > 1 or cols > 1):
                    cursor.currentTable().splitCell(rows, cols)
    
    def show_table_properties(self):
        """Show table properties dialog."""
        editor = self.get_current_editor()
        if editor:
            cursor = editor.textCursor()
            if cursor.currentTable():
                dialog = TablePropertiesDialog(cursor.currentTable(), self)
                if dialog.exec() == QDialog.Accepted:
                    # Apply table properties
                    pass
    
    # View methods
    def zoom_in(self):
        """Zoom in the document."""
        editor = self.get_current_editor()
        if editor:
            editor.zoom_in()
            self.update_status_bar()
    
    def zoom_out(self):
        """Zoom out the document."""
        editor = self.get_current_editor()
        if editor:
            editor.zoom_out()
            self.update_status_bar()
    
    def zoom_reset(self):
        """Reset zoom to 100%."""
        editor = self.get_current_editor()
        if editor:
            editor.zoom_reset()
            self.update_status_bar()
    
    def toggle_standard_toolbar(self, visible):
        """Toggle the standard toolbar visibility."""
        self.main_toolbar.setVisible(visible)
    
    def toggle_formatting_toolbar(self, visible):
        """Toggle the formatting toolbar visibility."""
        self.format_toolbar.setVisible(visible)
    
    def toggle_navigation_panel(self, visible):
        """Toggle the navigation panel visibility."""
        self.left_panel.setVisible(visible)
    
    def toggle_styles_panel(self, visible):
        """Toggle the styles panel visibility."""
        # Styles panel is a tab in the left panel
        if visible:
            self.left_tab_widget.setCurrentWidget(self.styles_panel)
        self.left_panel.setVisible(True)
    
    def toggle_document_map(self, visible):
        """Toggle the document map visibility."""
        self.right_panel.setVisible(visible)
    
    def toggle_comments_panel(self, visible):
        """Toggle the comments panel visibility."""
        # Comments panel is a tab in the right panel
        if visible:
            self.right_tab_widget.setCurrentWidget(self.comments_panel)
        self.right_panel.setVisible(True)
    
    # Edit methods
    def undo(self):
        """Undo the last action."""
        editor = self.current_editor()
        if editor and editor.document().isUndoAvailable():
            editor.undo()
            self.update_word_count()
    
    def redo(self):
        """Redo the last undone action."""
        editor = self.current_editor()
        if editor and editor.document().isRedoAvailable():
            editor.redo()
            self.update_word_count()
    
    def cut(self):
        """Cut selected text to clipboard."""
        editor = self.get_current_editor()
        if editor:
            editor.cut()
    
    def copy(self):
        """Copy selected text to clipboard."""
        editor = self.get_current_editor()
        if editor:
            editor.copy()
    
    def paste(self):
        """Paste text from clipboard."""
        editor = self.get_current_editor()
        if editor:
            editor.paste()
    
    def select_all(self):
        """Select all text in the document."""
        editor = self.get_current_editor()
        if editor:
            editor.select_all()
    
    def find(self):
        """Open find dialog."""
        dialog = FindReplaceDialog(self, find_only=True)
        dialog.show()
    
    def replace(self):
        """Open replace dialog."""
        dialog = FindReplaceDialog(self, find_only=False)
        dialog.show()
    
    def go_to(self):
        """Open go to dialog."""
        dialog = GoToDialog(self)
        if dialog.exec() == QDialog.Accepted:
            # Go to the specified location
            pass
    
    # Tools methods
    def show_word_count(self):
        """Show word count dialog."""
        editor = self.get_current_editor()
        if editor:
            stats = editor.word_count()
            dialog = WordCountDialog(stats, self)
            dialog.exec()
    
    def show_options(self):
        """Show options dialog."""
        dialog = OptionsDialog(self)
        if dialog.exec() == QDialog.Accepted:
            # Apply settings
            pass
    
    # Help methods
    def show_about(self):
        """Show about dialog."""
        dialog = AboutDialog(self)
        dialog.exec()
    
    # File methods
    def page_setup(self):
        """Open page setup dialog."""
        dialog = PageSetupDialog(self)
        if dialog.exec() == QDialog.Accepted:
            # Apply page settings
            pass
    
    def print_preview(self):
        """Open print preview dialog."""
        dialog = PrintPreviewDialog(self)
        dialog.exec()
    
    def print_document(self):
        """Print the current document."""
        # TODO: Implement printing
        QMessageBox.information(self, "Print", "Print functionality will be implemented in a future version.")
    
    # Settings
    def load_settings(self):
        """Load application settings."""
        # Load window geometry and state
        if self.settings.contains("geometry"):
            self.restoreGeometry(self.settings.value("geometry"))
        if self.settings.contains("windowState"):
            self.restoreState(self.settings.value("windowState"))
        
        # Load other settings
        # TODO: Load other application settings
    
    def save_settings(self):
        """Save application settings."""
        # Save window geometry and state
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())
        
        # Save other settings
        # TODO: Save other application settings

    def on_document_activated(self, document: Document):
        """Handle when a document is activated.
        
        Args:
            document: The activated document
        """
        self.current_document = document
        editor = self.current_editor()
        
        if editor:
            # Connect document change signals
            editor.undoAvailable.connect(self.actionUndo.setEnabled)
            editor.redoAvailable.connect(self.actionRedo.setEnabled)
            editor.textChanged.connect(self.on_text_changed)
            
            # Update undo/redo button states
            self.actionUndo.setEnabled(editor.document().isUndoAvailable())
            self.actionRedo.setEnabled(editor.document().isRedoAvailable())
            
            # Update word count
            self.update_word_count()

    def on_text_changed(self):
        """Handle text changes in the editor."""
        if self.current_document:
            self.current_document.modified = True
            self.current_document.modified_at = datetime.now()
            self.update_window_title()
            self.update_word_count()

    def update_window_title(self):
        """Update window title to match Microsoft Word style."""
        if hasattr(self, 'current_document') and self.current_document:
            doc_name = self.current_document.title or "Document1"
            if self.current_document.modified:
                self.setWindowTitle(f"{doc_name}* - PyWord")
            else:
                self.setWindowTitle(f"{doc_name} - PyWord")
        else:
            self.setWindowTitle("Document1 - PyWord")
