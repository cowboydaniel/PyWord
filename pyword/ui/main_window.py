from datetime import datetime
from pathlib import Path
from PySide6.QtWidgets import (QMainWindow, QVBoxLayout, QWidget, QSplitter, QTabWidget,
                             QStatusBar, QMenuBar, QMenu, QToolBar, QFileDialog, QMessageBox,
                             QDockWidget, QLabel, QSizePolicy, QApplication, QDialog,
                             QHBoxLayout, QPushButton, QTextEdit, QLineEdit, QListWidget, QTreeView,
                             QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QSpinBox,
                             QDoubleSpinBox, QCheckBox, QGroupBox, QFormLayout, QScrollArea,
                             QFrame, QToolButton, QStyle, QColorDialog, QFontDialog, QInputDialog,
                             QSplitterHandle, QSlider)
from PySide6.QtPrintSupport import QPrinter, QPageSetupDialog, QPrintDialog, QPrintPreviewDialog
from PySide6.QtCore import Qt, QSize, QSettings, QTimer, QUrl, QMimeData, Signal
from PySide6.QtGui import (QAction, QIcon, QFont, QTextCursor, QTextCharFormat, QTextListFormat,
                         QTextBlockFormat, QTextTableFormat, QTextFrameFormat, QTextLength,
                         QTextImageFormat, QPixmap, QImage, QPainter, QPalette, QColor,
                         QDesktopServices, QKeySequence, QTextDocument, QTextFrame, QTextTable,
                         QTextDocumentFragment, QTextBlock, QTextList, QTextFormat,
                         QShortcut)

from ..core.document import Document, DocumentManager
from ..core.editor import TextEditor
from ..core.page_setup import PageSetup, PageOrientation, PageMargins
from ..core.print_manager import PrintManager
from .document_manager_ui import DocumentManagerUI
from .toolbars import MainToolBar, FormatToolBar, TableToolBar, ReviewToolBar, ViewToolBar
from .panels import NavigationPanel, StylesPanel, DocumentMapPanel, CommentsPanel
from .ribbon import RibbonBar
from .theme_manager import ThemeManager, Theme
from .quick_access_toolbar import QuickAccessToolbar
from .ruler import HorizontalRuler, VerticalRuler
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
        self.current_document = None  # Initialize current document

        # Initialize and apply theme (Microsoft Word style)
        self.theme_manager = ThemeManager(self)
        self.theme_manager.apply_theme()

        # Create document manager UI (before setup_ui to avoid AttributeError)
        self.document_ui = DocumentManagerUI(self.document_manager, self)

        # Connect document UI signals
        self.document_ui.document_activated.connect(self.on_document_activated)
        self.document_ui.document_closed.connect(self.on_document_closed)
        self.document_ui.document_saved.connect(self.on_document_saved)

        # Setup UI
        self.setup_ui()

        # Load settings
        self.load_settings()
        
        # Connect text changes to update word count
        if self.current_editor():
            self.current_editor().textChanged.connect(self.update_word_count)
        
        # Setup zoom shortcuts
        self.setup_zoom_shortcuts()

        # Create undo/redo actions (needed for shortcuts and connections)
        self.setup_edit_actions()

        # Create a new document by default if none exists
        if not self.document_manager.documents:
            self.document_ui.new_document()
    
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
        word_count = len(text.split()) if text.strip() else 0

        # Update status bar (Microsoft Word style)
        self.page_word_label.setText(f"Page 1 of 1  {word_count} words")
    
    def update_zoom_display(self, zoom_level: int):
        """Update the zoom level display in the status bar."""
        self.current_zoom = zoom_level
        self.zoom_label.setText(f"{zoom_level}%")

        # Update zoom slider
        if hasattr(self, 'zoom_slider'):
            self.zoom_slider.blockSignals(True)  # Prevent recursive calls
            self.zoom_slider.setValue(zoom_level)
            self.zoom_slider.blockSignals(False)

        # Update rulers to reflect zoom level
        if hasattr(self, 'horizontal_ruler'):
            self.horizontal_ruler.set_zoom(zoom_level)
        if hasattr(self, 'vertical_ruler'):
            self.vertical_ruler.set_zoom(zoom_level)
    
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

    def setup_edit_actions(self):
        """Setup edit actions (undo/redo) for keyboard shortcuts."""
        # Undo action
        self.actionUndo = QAction("&Undo", self)
        self.actionUndo.setShortcut("Ctrl+Z")
        self.actionUndo.triggered.connect(self.undo)
        self.actionUndo.setEnabled(False)
        self.addAction(self.actionUndo)

        # Redo action
        self.actionRedo = QAction("&Redo", self)
        self.actionRedo.setShortcut("Ctrl+Y")
        self.actionRedo.triggered.connect(self.redo)
        self.actionRedo.setEnabled(False)
        self.addAction(self.actionRedo)
    
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

        # Create left panel (navigation, styles, etc.)
        self.setup_left_panel()

        # Create editor area (will contain document_ui's tab widget)
        self.setup_editor_area()
        
        # Create right panel (document map, comments, etc.)
        self.setup_right_panel()
        
        # Create status bar
        self.setup_status_bar()

        # Create menus (commented out for Word-like interface)
        # self.setup_menus()

        # Hide side panels by default (make them toggleable)
        self.left_panel.setVisible(False)
        self.right_panel.setVisible(False)

        # Update UI
        self.update_ui()
    
    def setup_toolbars(self):
        """Create and setup the ribbon interface."""
        # Create Quick Access Toolbar (above ribbon)
        self.quick_access_toolbar = QuickAccessToolbar(self)
        self.addToolBar(Qt.TopToolBarArea, self.quick_access_toolbar)

        # Connect Quick Access Toolbar commands
        self.quick_access_toolbar.connect_command('save', self.save_document)
        self.quick_access_toolbar.connect_command('undo', self.undo)
        self.quick_access_toolbar.connect_command('redo', self.redo)
        self.quick_access_toolbar.connect_command('print', self.print_document)

        # Create ribbon interface (Microsoft Word style)
        self.ribbon = RibbonBar(self)
        self.main_layout.addWidget(self.ribbon)

        # Create and add ribbon tabs (File tab first, like Word)
        file_tab = self.ribbon.create_file_tab()
        home_tab = self.ribbon.create_home_tab()
        insert_tab = self.ribbon.create_insert_tab()
        design_tab = self.ribbon.create_design_tab()
        layout_tab = self.ribbon.create_layout_tab()
        view_tab = self.ribbon.create_view_tab()

        self.ribbon.add_tab(file_tab)
        self.ribbon.add_tab(home_tab)
        self.ribbon.add_tab(insert_tab)
        self.ribbon.add_tab(design_tab)
        self.ribbon.add_tab(layout_tab)
        self.ribbon.add_tab(view_tab)

        # Connect ribbon actions to methods
        self.connect_ribbon_actions()

        # Connect File tab to show backstage view
        self.ribbon.file_tab_clicked.connect(self.ribbon.show_backstage)

    def connect_ribbon_actions(self):
        """Connect ribbon button actions to their respective methods."""
        # File tab connections
        if 'new' in self.ribbon.buttons:
            self.ribbon.buttons['new'].clicked.connect(self.new_document)
        if 'open' in self.ribbon.buttons:
            self.ribbon.buttons['open'].clicked.connect(self.open_document)
        if 'save' in self.ribbon.buttons:
            self.ribbon.buttons['save'].clicked.connect(self.save_document)
        if 'save_as' in self.ribbon.buttons:
            self.ribbon.buttons['save_as'].clicked.connect(self.save_document_as)
        if 'print' in self.ribbon.buttons:
            self.ribbon.buttons['print'].clicked.connect(self.print_document)
        if 'print_preview' in self.ribbon.buttons:
            self.ribbon.buttons['print_preview'].clicked.connect(self.print_preview)
        if 'export' in self.ribbon.buttons:
            self.ribbon.buttons['export'].clicked.connect(self.export_document)
        if 'export_pdf' in self.ribbon.buttons:
            self.ribbon.buttons['export_pdf'].clicked.connect(self.export_to_pdf)

        # Home tab - Clipboard group
        if 'cut' in self.ribbon.buttons:
            self.ribbon.buttons['cut'].clicked.connect(self.cut)
        if 'copy' in self.ribbon.buttons:
            self.ribbon.buttons['copy'].clicked.connect(self.copy)
        if 'paste' in self.ribbon.buttons:
            self.ribbon.buttons['paste'].clicked.connect(self.paste)
        if 'format_painter' in self.ribbon.buttons:
            self.ribbon.buttons['format_painter'].clicked.connect(self.format_painter)

        # Home tab - Font group
        if 'font_family' in self.ribbon.buttons:
            self.ribbon.buttons['font_family'].currentFontChanged.connect(self.on_font_family_changed)
        if 'font_size' in self.ribbon.buttons:
            self.ribbon.buttons['font_size'].valueChanged.connect(self.on_font_size_changed)
        if 'font_size_up' in self.ribbon.buttons:
            self.ribbon.buttons['font_size_up'].clicked.connect(self.increase_font_size)
        if 'font_size_down' in self.ribbon.buttons:
            self.ribbon.buttons['font_size_down'].clicked.connect(self.decrease_font_size)
        if 'bold' in self.ribbon.buttons:
            self.ribbon.buttons['bold'].clicked.connect(self.toggle_bold)
        if 'italic' in self.ribbon.buttons:
            self.ribbon.buttons['italic'].clicked.connect(self.toggle_italic)
        if 'underline' in self.ribbon.buttons:
            self.ribbon.buttons['underline'].clicked.connect(self.toggle_underline)
        if 'strikethrough' in self.ribbon.buttons:
            self.ribbon.buttons['strikethrough'].clicked.connect(self.toggle_strikethrough)
        if 'subscript' in self.ribbon.buttons:
            self.ribbon.buttons['subscript'].clicked.connect(self.toggle_subscript)
        if 'superscript' in self.ribbon.buttons:
            self.ribbon.buttons['superscript'].clicked.connect(self.toggle_superscript)
        if 'highlight_color' in self.ribbon.buttons:
            self.ribbon.buttons['highlight_color'].clicked.connect(self.select_highlight_color)
        if 'font_color' in self.ribbon.buttons:
            self.ribbon.buttons['font_color'].clicked.connect(self.select_font_color)
        if 'text_effects' in self.ribbon.buttons:
            self.ribbon.buttons['text_effects'].clicked.connect(self.text_effects)

        # Home tab - Paragraph group
        if 'bullets' in self.ribbon.buttons:
            self.ribbon.buttons['bullets'].clicked.connect(self.toggle_bullets)
        if 'numbering' in self.ribbon.buttons:
            self.ribbon.buttons['numbering'].clicked.connect(self.toggle_numbering)
        if 'multilevel_list' in self.ribbon.buttons:
            self.ribbon.buttons['multilevel_list'].clicked.connect(self.multilevel_list)
        if 'decrease_indent' in self.ribbon.buttons:
            self.ribbon.buttons['decrease_indent'].clicked.connect(self.decrease_indent)
        if 'increase_indent' in self.ribbon.buttons:
            self.ribbon.buttons['increase_indent'].clicked.connect(self.increase_indent)
        if 'align_left' in self.ribbon.buttons:
            self.ribbon.buttons['align_left'].clicked.connect(self.align_left)
        if 'align_center' in self.ribbon.buttons:
            self.ribbon.buttons['align_center'].clicked.connect(self.align_center)
        if 'align_right' in self.ribbon.buttons:
            self.ribbon.buttons['align_right'].clicked.connect(self.align_right)
        if 'align_justify' in self.ribbon.buttons:
            self.ribbon.buttons['align_justify'].clicked.connect(self.align_justify)
        if 'sort' in self.ribbon.buttons:
            self.ribbon.buttons['sort'].clicked.connect(self.sort_text)
        if 'show_hide' in self.ribbon.buttons:
            self.ribbon.buttons['show_hide'].clicked.connect(self.show_hide_formatting)
        if 'line_spacing' in self.ribbon.buttons:
            self.ribbon.buttons['line_spacing'].clicked.connect(self.line_spacing)
        if 'shading' in self.ribbon.buttons:
            self.ribbon.buttons['shading'].clicked.connect(self.paragraph_shading)
        if 'borders' in self.ribbon.buttons:
            self.ribbon.buttons['borders'].clicked.connect(self.paragraph_borders)

        # Home tab - Editing group
        if 'find' in self.ribbon.buttons:
            self.ribbon.buttons['find'].clicked.connect(self.find)
        if 'replace' in self.ribbon.buttons:
            self.ribbon.buttons['replace'].clicked.connect(self.replace)
        if 'select_all' in self.ribbon.buttons:
            self.ribbon.buttons['select_all'].clicked.connect(self.select_all)

        # Insert tab connections
        if 'insert_table' in self.ribbon.buttons:
            self.ribbon.buttons['insert_table'].clicked.connect(self.insert_table)
        if 'insert_picture' in self.ribbon.buttons:
            self.ribbon.buttons['insert_picture'].clicked.connect(self.insert_image)
        if 'insert_shapes' in self.ribbon.buttons:
            self.ribbon.buttons['insert_shapes'].clicked.connect(self.insert_shapes)
        if 'insert_chart' in self.ribbon.buttons:
            self.ribbon.buttons['insert_chart'].clicked.connect(self.insert_chart)
        if 'insert_smartart' in self.ribbon.buttons:
            self.ribbon.buttons['insert_smartart'].clicked.connect(self.insert_smartart)
        if 'insert_link' in self.ribbon.buttons:
            self.ribbon.buttons['insert_link'].clicked.connect(self.insert_hyperlink)
        if 'insert_bookmark' in self.ribbon.buttons:
            self.ribbon.buttons['insert_bookmark'].clicked.connect(self.insert_bookmark)
        if 'insert_header' in self.ribbon.buttons:
            self.ribbon.buttons['insert_header'].clicked.connect(self.insert_header)
        if 'insert_footer' in self.ribbon.buttons:
            self.ribbon.buttons['insert_footer'].clicked.connect(self.insert_footer)
        if 'insert_page_number' in self.ribbon.buttons:
            self.ribbon.buttons['insert_page_number'].clicked.connect(self.insert_page_number)
        if 'insert_equation' in self.ribbon.buttons:
            self.ribbon.buttons['insert_equation'].clicked.connect(self.insert_equation)
        if 'insert_symbol' in self.ribbon.buttons:
            self.ribbon.buttons['insert_symbol'].clicked.connect(self.insert_symbol)

        # Design tab connections
        if 'design_themes' in self.ribbon.buttons:
            self.ribbon.buttons['design_themes'].clicked.connect(self.design_themes)
        if 'design_colors' in self.ribbon.buttons:
            self.ribbon.buttons['design_colors'].clicked.connect(self.design_colors)
        if 'design_fonts' in self.ribbon.buttons:
            self.ribbon.buttons['design_fonts'].clicked.connect(self.design_fonts)
        if 'design_watermark' in self.ribbon.buttons:
            self.ribbon.buttons['design_watermark'].clicked.connect(self.design_watermark)
        if 'design_page_color' in self.ribbon.buttons:
            self.ribbon.buttons['design_page_color'].clicked.connect(self.design_page_color)
        if 'design_page_borders' in self.ribbon.buttons:
            self.ribbon.buttons['design_page_borders'].clicked.connect(self.design_page_borders)

        # Layout tab connections
        if 'layout_margins' in self.ribbon.buttons:
            self.ribbon.buttons['layout_margins'].clicked.connect(self.layout_margins)
        if 'layout_orientation' in self.ribbon.buttons:
            self.ribbon.buttons['layout_orientation'].clicked.connect(self.layout_orientation)
        if 'layout_size' in self.ribbon.buttons:
            self.ribbon.buttons['layout_size'].clicked.connect(self.layout_size)
        if 'layout_columns' in self.ribbon.buttons:
            self.ribbon.buttons['layout_columns'].clicked.connect(self.layout_columns)
        if 'layout_position' in self.ribbon.buttons:
            self.ribbon.buttons['layout_position'].clicked.connect(self.layout_position)
        if 'layout_wrap_text' in self.ribbon.buttons:
            self.ribbon.buttons['layout_wrap_text'].clicked.connect(self.layout_wrap_text)
        if 'layout_align' in self.ribbon.buttons:
            self.ribbon.buttons['layout_align'].clicked.connect(self.layout_align)

        # View tab connections
        if 'view_print_layout' in self.ribbon.buttons:
            self.ribbon.buttons['view_print_layout'].clicked.connect(self.view_print_layout)
        if 'view_web_layout' in self.ribbon.buttons:
            self.ribbon.buttons['view_web_layout'].clicked.connect(self.view_web_layout)
        if 'view_draft' in self.ribbon.buttons:
            self.ribbon.buttons['view_draft'].clicked.connect(self.view_draft)
        if 'view_zoom' in self.ribbon.buttons:
            self.ribbon.buttons['view_zoom'].clicked.connect(self.view_zoom_dialog)
        if 'view_zoom_100' in self.ribbon.buttons:
            self.ribbon.buttons['view_zoom_100'].clicked.connect(self.zoom_reset)
        if 'view_page_width' in self.ribbon.buttons:
            self.ribbon.buttons['view_page_width'].clicked.connect(self.view_page_width)
        if 'view_new_window' in self.ribbon.buttons:
            self.ribbon.buttons['view_new_window'].clicked.connect(self.view_new_window)
        if 'view_split' in self.ribbon.buttons:
            self.ribbon.buttons['view_split'].clicked.connect(self.view_split)
        if 'view_side_by_side' in self.ribbon.buttons:
            self.ribbon.buttons['view_side_by_side'].clicked.connect(self.view_side_by_side)

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

        # Set gray background for the editor area (Microsoft Word style)
        self.editor_area.setStyleSheet("""
            QWidget {
                background-color: #D3D3D3;
            }
        """)

        editor_layout = QVBoxLayout(self.editor_area)
        editor_layout.setContentsMargins(0, 0, 0, 0)
        editor_layout.setSpacing(0)

        # A4 page size at 96 DPI: 21cm x 29.7cm = 794px x 1122px
        self.page_width_px = 794
        self.page_height_px = 1122

        # Horizontal ruler - FULL WIDTH spanning workspace
        # It will calculate offset internally to align 0 with page edge
        self.horizontal_ruler = HorizontalRuler(self.editor_area)
        self.horizontal_ruler.page_width_px = self.page_width_px
        editor_layout.addWidget(self.horizontal_ruler)

        # Content row: vertical ruler + page
        content_row_layout = QHBoxLayout()
        content_row_layout.setContentsMargins(0, 0, 0, 0)
        content_row_layout.setSpacing(0)

        # Vertical ruler - anchored to left edge, FULL HEIGHT
        # It will calculate offset internally to align 0 with page edge
        self.vertical_ruler = VerticalRuler(self.editor_area)
        self.vertical_ruler.page_height_px = self.page_height_px
        content_row_layout.addWidget(self.vertical_ruler, 0, Qt.AlignLeft)

        # Page container - takes remaining space, centers the page
        page_container = QWidget()
        page_container.setStyleSheet("background-color: #D3D3D3;")
        page_container_layout = QHBoxLayout(page_container)
        page_container_layout.setContentsMargins(0, 0, 0, 0)
        page_container_layout.setSpacing(0)

        # Add left stretch to center the page
        page_container_layout.addStretch()

        # Use document_ui's tab widget (contains the actual text editors)
        self.tab_widget = self.document_ui.tab_widget

        # Style the tab widget to have gray background with white page (Microsoft Word style)
        # Set to A4 size at 100% zoom (21cm = 794px at 96 DPI)
        self.tab_widget.setFixedWidth(self.page_width_px)
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: none;
                background-color: #D3D3D3;
            }
            QTextEdit {
                background-color: white;
                border: 1px solid #A0A0A0;
                margin: 20px;
            }
        """)

        page_container_layout.addWidget(self.tab_widget)

        # Add right stretch to center the page
        page_container_layout.addStretch()

        # Add page container to the content row (takes remaining width)
        content_row_layout.addWidget(page_container, 1)

        editor_layout.addLayout(content_row_layout)

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
        """Setup the status bar (Microsoft Word style)."""
        # Create status bar
        self.status_bar = QStatusBar()
        self.status_bar.setStyleSheet("""
            QStatusBar {
                background-color: #F3F2F1;
                border-top: 1px solid #D2D0CE;
                padding: 2px 8px;
            }
            QLabel {
                color: #323130;
                font-size: 12px;
                padding: 0 4px;
            }
        """)
        self.setStatusBar(self.status_bar)

        # Left side: Page and word count
        self.page_word_label = QLabel("Page 1 of 1  0 words")
        self.status_bar.addWidget(self.page_word_label)

        # Add spacer
        self.status_bar.addWidget(QLabel(""), 1)  # Stretch

        # Right side: Zoom controls (Word style)
        zoom_widget = QWidget()
        zoom_layout = QHBoxLayout(zoom_widget)
        zoom_layout.setContentsMargins(0, 0, 0, 0)
        zoom_layout.setSpacing(4)

        # Zoom out button
        zoom_out_btn = QPushButton("-")
        zoom_out_btn.setFixedSize(20, 20)
        zoom_out_btn.setStyleSheet("""
            QPushButton {
                border: 1px solid #D2D0CE;
                border-radius: 2px;
                background: white;
                color: #323130;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #F3F2F1;
            }
            QPushButton:pressed {
                background-color: #EDEBE9;
            }
        """)
        zoom_out_btn.clicked.connect(self.zoom_out)
        zoom_layout.addWidget(zoom_out_btn)

        # Zoom slider
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setMinimum(25)
        self.zoom_slider.setMaximum(400)
        self.zoom_slider.setValue(100)
        self.zoom_slider.setFixedWidth(100)
        self.zoom_slider.setStyleSheet("""
            QSlider::groove:horizontal {
                border: 1px solid #D2D0CE;
                height: 4px;
                background: white;
                margin: 0px;
                border-radius: 2px;
            }
            QSlider::handle:horizontal {
                background: #0078D4;
                border: 1px solid #0078D4;
                width: 12px;
                height: 12px;
                margin: -5px 0;
                border-radius: 6px;
            }
            QSlider::handle:horizontal:hover {
                background: #106EBE;
                border: 1px solid #106EBE;
            }
        """)
        self.zoom_slider.valueChanged.connect(self.on_zoom_slider_changed)
        zoom_layout.addWidget(self.zoom_slider)

        # Zoom in button
        zoom_in_btn = QPushButton("+")
        zoom_in_btn.setFixedSize(20, 20)
        zoom_in_btn.setStyleSheet("""
            QPushButton {
                border: 1px solid #D2D0CE;
                border-radius: 2px;
                background: white;
                color: #323130;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #F3F2F1;
            }
            QPushButton:pressed {
                background-color: #EDEBE9;
            }
        """)
        zoom_in_btn.clicked.connect(self.zoom_in)
        zoom_layout.addWidget(zoom_in_btn)

        # Zoom percentage label
        self.zoom_label = QLabel("100%")
        self.zoom_label.setMinimumWidth(40)
        zoom_layout.addWidget(self.zoom_label)

        self.status_bar.addPermanentWidget(zoom_widget)

        # Connect zoom signal
        self.zoom_changed.connect(self.update_zoom_display)

    def on_zoom_slider_changed(self, value):
        """Handle zoom slider value change."""
        if self.current_editor():
            # Set zoom factor on editor
            self.current_editor().zoom_factor = value / 100.0
            self.current_editor().update_zoom()
            self.current_zoom = value
    
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

        # Add existing undo/redo actions (created in setup_edit_actions)
        edit_menu.addAction(self.actionUndo)
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
        table_properties_action.triggered.connect(self.show_table_properties)
        self.table_menu.addAction(table_properties_action)

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
            self.status_bar.clearMessage()
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
        self.status_bar.showMessage(" | ".join(status_text))

    # Document management
    def new_document(self):
        """Create a new document."""
        self.document_ui.new_document()

    def open_document(self):
        """Open an existing document."""
        self.document_ui.open_document_dialog()

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

    def current_editor(self):
        """Get the current active editor."""
        return self.document_ui.tab_widget.current_editor()

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
        editor = self.current_editor()
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
        editor = self.current_editor()
        if editor:
            font, ok = QFontDialog.getFont(editor.currentFont(), self, "Select Font")
            if ok:
                editor.setCurrentFont(font)
    
    def format_paragraph(self):
        """Open paragraph formatting dialog."""
        editor = self.current_editor()
        if editor:
            # Get current paragraph format from the editor
            cursor = editor.textCursor()
            initial_format = cursor.blockFormat() if cursor else None
            dialog = ParagraphDialog(self, initial_format)
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
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            cursor.insertHtml("<hr style='page-break-before:always; margin:0;' />")
    
    def insert_table(self):
        """Insert a table at the current cursor position."""
        editor = self.current_editor()
        if editor:
            dialog = InsertTableDialog(self)
            if dialog.exec() == QDialog.Accepted:
                rows = dialog.get_rows()
                cols = dialog.get_columns()
                editor.insert_table(rows, cols)
    
    def insert_image(self):
        """Insert an image at the current cursor position."""
        editor = self.current_editor()
        if editor:
            dialog = InsertImageDialog(self)
            if dialog.exec() == QDialog.Accepted:
                image_path = dialog.get_image_path()
                if image_path:
                    editor.insert_image(image_path)
    
    def insert_hyperlink(self):
        """Insert a hyperlink at the current cursor position."""
        editor = self.current_editor()
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
                editor = self.current_editor()
                if editor:
                    editor.insertPlainText(symbol)
    
    # Table operations
    def delete_table(self):
        """Delete the current table."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            table = cursor.currentTable()
            if table:
                # Get the table's frame and remove it entirely
                cursor.movePosition(QTextCursor.Start)
                cursor = table.firstCursorPosition()
                cursor.movePosition(QTextCursor.PreviousCharacter)
                cursor.movePosition(QTextCursor.NextCharacter, QTextCursor.KeepAnchor,
                                  table.lastCursorPosition().position() - cursor.position() + 1)
                cursor.removeSelectedText()
    
    def insert_table_row(self, above):
        """Insert a row above or below the current row."""
        editor = self.current_editor()
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
        editor = self.current_editor()
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
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            table = cursor.currentTable()
            if table:
                table.removeRows(cursor.currentRow(), 1)
    
    def delete_table_column(self):
        """Delete the current column."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            table = cursor.currentTable()
            if table:
                table.removeColumns(cursor.columnNumber(), 1)
    
    def merge_table_cells(self):
        """Merge selected table cells."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            if cursor.hasSelection() and cursor.currentTable():
                cursor.currentTable().mergeCells(cursor)
    
    def split_table_cells(self):
        """Split the current table cell."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            if cursor.currentTable():
                # Get number of rows
                rows, ok1 = QInputDialog.getInt(self, "Split Cells", "Number of rows:", 1, 1, 100, 1)
                if not ok1:
                    return

                # Get number of columns
                cols, ok2 = QInputDialog.getInt(self, "Split Cells", "Number of columns:", 1, 1, 100, 1)
                if not ok2:
                    return

                if rows > 1 or cols > 1:
                    cursor.currentTable().splitCell(rows, cols)
    
    def show_table_properties(self):
        """Show table properties dialog."""
        editor = self.current_editor()
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
        editor = self.current_editor()
        if editor:
            editor.zoom_in()
            self.update_status_bar()
    
    def zoom_out(self):
        """Zoom out the document."""
        editor = self.current_editor()
        if editor:
            editor.zoom_out()
            self.update_status_bar()
    
    def zoom_reset(self):
        """Reset zoom to 100%."""
        editor = self.current_editor()
        if editor:
            editor.zoom_reset()
            self.update_status_bar()
    
    def toggle_standard_toolbar(self, visible):
        """Toggle the ribbon visibility."""
        if hasattr(self, 'ribbon'):
            self.ribbon.setVisible(visible)

    def toggle_formatting_toolbar(self, visible):
        """Toggle the formatting toolbar visibility.

        Note: In this implementation, both standard and formatting tools
        are part of the same ribbon, so this toggles the ribbon visibility.
        """
        # Delegate to toggle_standard_toolbar since they control the same ribbon
        self.toggle_standard_toolbar(visible)
    
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
        editor = self.current_editor()
        if editor:
            editor.cut()
    
    def copy(self):
        """Copy selected text to clipboard."""
        editor = self.current_editor()
        if editor:
            editor.copy()
    
    def paste(self):
        """Paste text from clipboard."""
        editor = self.current_editor()
        if editor:
            editor.paste()
    
    def select_all(self):
        """Select all text in the document."""
        editor = self.current_editor()
        if editor:
            editor.select_all()

    # Font formatting methods
    def on_font_family_changed(self, font):
        """Handle font family change."""
        editor = self.current_editor()
        if editor:
            editor.set_font_family(font.family())

    def on_font_size_changed(self, size):
        """Handle font size change."""
        editor = self.current_editor()
        if editor:
            editor.set_font_size(size)

    def increase_font_size(self):
        """Increase font size by 1 point."""
        if 'font_size' in self.ribbon.buttons:
            current = self.ribbon.buttons['font_size'].value()
            self.ribbon.buttons['font_size'].setValue(current + 1)

    def decrease_font_size(self):
        """Decrease font size by 1 point."""
        if 'font_size' in self.ribbon.buttons:
            current = self.ribbon.buttons['font_size'].value()
            self.ribbon.buttons['font_size'].setValue(max(8, current - 1))

    def toggle_bold(self):
        """Toggle bold formatting."""
        editor = self.current_editor()
        if editor:
            editor.text_bold()

    def toggle_italic(self):
        """Toggle italic formatting."""
        editor = self.current_editor()
        if editor:
            editor.text_italic()

    def toggle_underline(self):
        """Toggle underline formatting."""
        editor = self.current_editor()
        if editor:
            editor.text_underline()

    def toggle_strikethrough(self):
        """Toggle strikethrough formatting."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            fmt = cursor.charFormat()
            fmt.setFontStrikeOut(not fmt.fontStrikeOut())
            if not cursor.hasSelection():
                cursor.select(QTextCursor.WordUnderCursor)
            cursor.mergeCharFormat(fmt)

    def toggle_subscript(self):
        """Toggle subscript formatting."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            fmt = cursor.charFormat()
            if fmt.verticalAlignment() == QTextCharFormat.AlignSubScript:
                fmt.setVerticalAlignment(QTextCharFormat.AlignNormal)
            else:
                fmt.setVerticalAlignment(QTextCharFormat.AlignSubScript)
            if not cursor.hasSelection():
                cursor.select(QTextCursor.WordUnderCursor)
            cursor.mergeCharFormat(fmt)

    def toggle_superscript(self):
        """Toggle superscript formatting."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            fmt = cursor.charFormat()
            if fmt.verticalAlignment() == QTextCharFormat.AlignSuperScript:
                fmt.setVerticalAlignment(QTextCharFormat.AlignNormal)
            else:
                fmt.setVerticalAlignment(QTextCharFormat.AlignSuperScript)
            if not cursor.hasSelection():
                cursor.select(QTextCursor.WordUnderCursor)
            cursor.mergeCharFormat(fmt)

    def select_highlight_color(self):
        """Select highlight color."""
        editor = self.current_editor()
        if editor:
            color = QColorDialog.getColor(Qt.yellow, self, "Select Highlight Color")
            if color.isValid():
                editor.set_highlight_color(color)

    def select_font_color(self):
        """Select font color."""
        editor = self.current_editor()
        if editor:
            color = QColorDialog.getColor(Qt.black, self, "Select Font Color")
            if color.isValid():
                editor.set_text_color(color)

    # Paragraph formatting methods
    def toggle_bullets(self):
        """Toggle bullet list."""
        editor = self.current_editor()
        if editor:
            editor.insert_bullet_list()

    def toggle_numbering(self):
        """Toggle numbered list."""
        editor = self.current_editor()
        if editor:
            editor.insert_numbered_list()

    def decrease_indent(self):
        """Decrease indent."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            fmt = cursor.blockFormat()
            indent = fmt.indent()
            if indent > 0:
                fmt.setIndent(indent - 1)
                cursor.setBlockFormat(fmt)

    def increase_indent(self):
        """Increase indent."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            fmt = cursor.blockFormat()
            fmt.setIndent(fmt.indent() + 1)
            cursor.setBlockFormat(fmt)

    def align_left(self):
        """Align text left."""
        editor = self.current_editor()
        if editor:
            editor.set_alignment(Qt.AlignLeft)

    def align_center(self):
        """Align text center."""
        editor = self.current_editor()
        if editor:
            editor.set_alignment(Qt.AlignCenter)

    def align_right(self):
        """Align text right."""
        editor = self.current_editor()
        if editor:
            editor.set_alignment(Qt.AlignRight)

    def align_justify(self):
        """Justify text."""
        editor = self.current_editor()
        if editor:
            editor.set_alignment(Qt.AlignJustify)
    
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
        editor = self.current_editor()
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
    
    def export_document(self):
        """Export document to a different format."""
        # Get current document from document_ui
        current_doc = self.document_ui.current_document if hasattr(self, 'document_ui') else None

        if not current_doc and not self.current_editor():
            QMessageBox.warning(self, "No Document", "Please open a document first.")
            return

        # Use current_doc if available, otherwise fall back to self.current_document
        doc = current_doc or self.current_document

        # Show save dialog with various export formats
        file_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Export Document",
            doc.file_path if doc and hasattr(doc, 'file_path') else "",
            "PDF Files (*.pdf);;"
            "Word Documents (*.docx);;"
            "OpenDocument Text (*.odt);;"
            "HTML Files (*.html);;"
            "Rich Text Files (*.rtf);;"
            "Text Files (*.txt);;"
            "All Files (*.*)"
        )

        if file_path:
            try:
                if doc and doc.save(file_path):
                    QMessageBox.information(self, "Export Successful", f"Document exported to {file_path}")
                else:
                    QMessageBox.warning(self, "Export Failed", "Failed to export document.")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Error exporting document: {str(e)}")

    def export_to_pdf(self):
        """Export document directly to PDF format."""
        # Get current document from document_ui
        current_doc = self.document_ui.current_document if hasattr(self, 'document_ui') else None

        if not current_doc and not self.current_editor():
            QMessageBox.warning(self, "No Document", "Please open a document first.")
            return

        # Use current_doc if available, otherwise fall back to self.current_document
        doc = current_doc or self.current_document

        # Suggest PDF filename based on current document
        suggested_name = ""
        if doc and hasattr(doc, 'file_path') and doc.file_path:
            suggested_name = str(Path(doc.file_path).with_suffix('.pdf'))
        elif doc and hasattr(doc, 'title'):
            suggested_name = f"{doc.title}.pdf"
        else:
            suggested_name = "document.pdf"

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export to PDF",
            suggested_name,
            "PDF Files (*.pdf)"
        )

        if file_path:
            try:
                if doc and doc.save(file_path):
                    QMessageBox.information(self, "Export Successful", f"Document exported to PDF: {file_path}")
                else:
                    QMessageBox.warning(self, "Export Failed", "Failed to export document to PDF.")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Error exporting to PDF: {str(e)}")

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

    def on_document_closed(self, document: Document):
        """Handle when a document is closed.

        Args:
            document: The closed document
        """
        # Update current document if the closed one was active
        if self.current_document == document:
            self.current_document = None

        # Update UI
        self.update_ui()
        self.update_window_title()

    def on_document_saved(self, document: Document, file_path: str):
        """Handle when a document is saved.

        Args:
            document: The saved document
            file_path: Path where document was saved
        """
        # Update window title to remove modified indicator
        self.update_window_title()

        # Update status bar
        self.update_status_bar()

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

    # Home tab - Additional formatting methods
    def format_painter(self):
        """Copy formatting from selected text to apply elsewhere."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            if cursor.hasSelection():
                # Store the format from selected text
                self.stored_format = cursor.charFormat()
                QMessageBox.information(self, "Format Painter", "Format copied! Select text to apply the formatting.")
            elif hasattr(self, 'stored_format'):
                # Apply stored format
                if not cursor.hasSelection():
                    cursor.select(QTextCursor.WordUnderCursor)
                cursor.mergeCharFormat(self.stored_format)
                delattr(self, 'stored_format')
            else:
                QMessageBox.information(self, "Format Painter", "Please select text to copy formatting from.")

    def text_effects(self):
        """Apply text effects (shadow, outline, etc.)."""
        editor = self.current_editor()
        if editor:
            # Simple implementation: show available effects
            items = ["Shadow", "Outline", "Glow", "Reflection"]
            item, ok = QInputDialog.getItem(self, "Text Effects", "Select effect:", items, 0, False)
            if ok and item:
                QMessageBox.information(self, "Text Effects", f"Text effect '{item}' would be applied here.\nThis is a placeholder implementation.")

    def multilevel_list(self):
        """Insert a multilevel list."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            # Create a multilevel list format
            list_format = QTextListFormat()
            list_format.setStyle(QTextListFormat.ListDecimal)
            list_format.setIndent(2)
            cursor.createList(list_format)

    def sort_text(self):
        """Sort selected paragraphs alphabetically."""
        editor = self.current_editor()
        if editor:
            cursor = editor.textCursor()
            if cursor.hasSelection():
                # Get selected text
                text = cursor.selectedText()
                # Split by paragraph separator
                lines = text.split('\u2029')
                # Sort lines
                lines.sort()
                # Replace selection with sorted text
                cursor.insertText('\u2029'.join(lines))
            else:
                QMessageBox.information(self, "Sort", "Please select text to sort.")

    def show_hide_formatting(self):
        """Toggle visibility of formatting marks."""
        editor = self.current_editor()
        if editor:
            # Toggle a flag to show/hide formatting marks
            if not hasattr(editor, 'show_formatting_marks'):
                editor.show_formatting_marks = False
            editor.show_formatting_marks = not editor.show_formatting_marks

            if editor.show_formatting_marks:
                QMessageBox.information(self, "Show/Hide", "Formatting marks are now visible.")
            else:
                QMessageBox.information(self, "Show/Hide", "Formatting marks are now hidden.")

    def line_spacing(self):
        """Set line spacing for selected paragraphs."""
        editor = self.current_editor()
        if editor:
            items = ["Single", "1.15", "1.5", "Double", "2.5", "3.0"]
            item, ok = QInputDialog.getItem(self, "Line Spacing", "Select line spacing:", items, 0, False)
            if ok and item:
                cursor = editor.textCursor()
                block_format = cursor.blockFormat()

                # Map selection to spacing value
                spacing_map = {"Single": 100, "1.15": 115, "1.5": 150, "Double": 200, "2.5": 250, "3.0": 300}
                spacing = spacing_map.get(item, 100)

                block_format.setLineHeight(spacing, QTextBlockFormat.ProportionalHeight)
                cursor.setBlockFormat(block_format)

    def paragraph_shading(self):
        """Apply background color to paragraph."""
        editor = self.current_editor()
        if editor:
            color = QColorDialog.getColor(Qt.yellow, self, "Select Shading Color")
            if color.isValid():
                cursor = editor.textCursor()
                block_format = cursor.blockFormat()
                block_format.setBackground(color)
                cursor.setBlockFormat(block_format)

    def paragraph_borders(self):
        """Apply borders to paragraph."""
        editor = self.current_editor()
        if editor:
            dialog = BorderAndShadingDialog(self)
            if dialog.exec() == QDialog.Accepted:
                QMessageBox.information(self, "Borders", "Paragraph borders would be applied here.")

    # Insert tab - Additional methods
    def insert_shapes(self):
        """Insert shapes into document."""
        QMessageBox.information(self, "Insert Shapes", "Shape insertion would open a shape picker dialog.\nThis is a placeholder implementation.")

    def insert_chart(self):
        """Insert a chart into document."""
        QMessageBox.information(self, "Insert Chart", "Chart insertion would open a chart creation dialog.\nThis is a placeholder implementation.")

    def insert_smartart(self):
        """Insert SmartArt graphic."""
        QMessageBox.information(self, "Insert SmartArt", "SmartArt insertion would open a SmartArt picker dialog.\nThis is a placeholder implementation.")

    def insert_bookmark(self):
        """Insert a bookmark at current position."""
        editor = self.current_editor()
        if editor:
            name, ok = QInputDialog.getText(self, "Insert Bookmark", "Bookmark name:")
            if ok and name:
                cursor = editor.textCursor()
                # In a full implementation, would store bookmark position
                QMessageBox.information(self, "Bookmark", f"Bookmark '{name}' inserted at current position.")

    def insert_header(self):
        """Edit document header."""
        QMessageBox.information(self, "Insert Header", "Header editing would open header editor.\nThis is a placeholder implementation.")

    def insert_footer(self):
        """Edit document footer."""
        QMessageBox.information(self, "Insert Footer", "Footer editing would open footer editor.\nThis is a placeholder implementation.")

    def insert_page_number(self):
        """Insert page number."""
        editor = self.current_editor()
        if editor:
            items = ["Top of Page", "Bottom of Page", "Page Margins"]
            item, ok = QInputDialog.getItem(self, "Page Number", "Select position:", items, 0, False)
            if ok and item:
                QMessageBox.information(self, "Page Number", f"Page number would be inserted at: {item}")

    def insert_equation(self):
        """Insert mathematical equation."""
        editor = self.current_editor()
        if editor:
            equation, ok = QInputDialog.getText(self, "Insert Equation", "Enter equation (LaTeX format):")
            if ok and equation:
                cursor = editor.textCursor()
                cursor.insertText(f"[Equation: {equation}]")

    # Design tab methods
    def design_themes(self):
        """Apply document theme."""
        items = ["Office", "Facet", "Integral", "Ion", "Organic", "Retrospect", "Slice"]
        item, ok = QInputDialog.getItem(self, "Themes", "Select theme:", items, 0, False)
        if ok and item:
            QMessageBox.information(self, "Theme", f"Theme '{item}' would be applied to document.")

    def design_colors(self):
        """Change theme colors."""
        items = ["Office", "Grayscale", "Blue Warm", "Blue", "Green", "Orange", "Red"]
        item, ok = QInputDialog.getItem(self, "Theme Colors", "Select color scheme:", items, 0, False)
        if ok and item:
            QMessageBox.information(self, "Theme Colors", f"Color scheme '{item}' would be applied.")

    def design_fonts(self):
        """Change theme fonts."""
        items = ["Office", "Calibri", "Arial", "Times New Roman", "Georgia", "Cambria"]
        item, ok = QInputDialog.getItem(self, "Theme Fonts", "Select font theme:", items, 0, False)
        if ok and item:
            QMessageBox.information(self, "Theme Fonts", f"Font theme '{item}' would be applied.")

    def design_watermark(self):
        """Add watermark to document."""
        items = ["Confidential", "Draft", "Sample", "Custom Text", "Remove Watermark"]
        item, ok = QInputDialog.getItem(self, "Watermark", "Select watermark:", items, 0, False)
        if ok and item:
            if item == "Custom Text":
                text, ok2 = QInputDialog.getText(self, "Custom Watermark", "Enter watermark text:")
                if ok2 and text:
                    QMessageBox.information(self, "Watermark", f"Watermark '{text}' would be added.")
            else:
                QMessageBox.information(self, "Watermark", f"Watermark '{item}' would be applied.")

    def design_page_color(self):
        """Set page background color."""
        color = QColorDialog.getColor(Qt.white, self, "Select Page Color")
        if color.isValid():
            editor = self.current_editor()
            if editor:
                palette = editor.palette()
                palette.setColor(QPalette.Base, color)
                editor.setPalette(palette)

    def design_page_borders(self):
        """Add borders to page."""
        dialog = BorderAndShadingDialog(self)
        if dialog.exec() == QDialog.Accepted:
            QMessageBox.information(self, "Page Borders", "Page borders would be applied.")

    # Layout tab methods
    def layout_margins(self):
        """Set page margins."""
        items = ["Normal", "Narrow", "Moderate", "Wide", "Mirrored", "Custom"]
        item, ok = QInputDialog.getItem(self, "Margins", "Select margins:", items, 0, False)
        if ok and item:
            if item == "Custom":
                self.page_setup()
            else:
                QMessageBox.information(self, "Margins", f"Margins set to '{item}'.")

    def layout_orientation(self):
        """Set page orientation."""
        items = ["Portrait", "Landscape"]
        item, ok = QInputDialog.getItem(self, "Orientation", "Select orientation:", items, 0, False)
        if ok and item:
            if self.current_document:
                from ..core.page_setup import PageOrientation
                self.current_document.page_setup.orientation = PageOrientation.PORTRAIT if item == "Portrait" else PageOrientation.LANDSCAPE
                QMessageBox.information(self, "Orientation", f"Page orientation set to {item}.")

    def layout_size(self):
        """Set page size."""
        items = ["Letter", "A4", "Legal", "Executive", "A5", "Custom"]
        item, ok = QInputDialog.getItem(self, "Page Size", "Select size:", items, 0, False)
        if ok and item:
            if item == "Custom":
                self.page_setup()
            else:
                QMessageBox.information(self, "Page Size", f"Page size set to '{item}'.")

    def layout_columns(self):
        """Set number of columns."""
        dialog = ColumnsDialog(self)
        if dialog.exec() == QDialog.Accepted:
            QMessageBox.information(self, "Columns", "Column layout would be applied.")

    def layout_position(self):
        """Set object position."""
        QMessageBox.information(self, "Position", "Object positioning controls would appear here.\nSelect an object first.")

    def layout_wrap_text(self):
        """Set text wrapping around objects."""
        items = ["Square", "Tight", "Through", "Top and Bottom", "Behind Text", "In Front of Text"]
        item, ok = QInputDialog.getItem(self, "Wrap Text", "Select wrapping style:", items, 0, False)
        if ok and item:
            QMessageBox.information(self, "Wrap Text", f"Text wrapping set to '{item}'.\nSelect an object first.")

    def layout_align(self):
        """Align selected objects."""
        items = ["Align Left", "Align Center", "Align Right", "Align Top", "Align Middle", "Align Bottom"]
        item, ok = QInputDialog.getItem(self, "Align", "Select alignment:", items, 0, False)
        if ok and item:
            QMessageBox.information(self, "Align", f"Objects would be aligned: {item}.\nSelect objects first.")

    # View tab methods
    def view_print_layout(self):
        """Switch to print layout view."""
        QMessageBox.information(self, "Print Layout", "Switched to Print Layout view.\n(Current view)")

    def view_web_layout(self):
        """Switch to web layout view."""
        QMessageBox.information(self, "Web Layout", "Web Layout view would be activated.\nThis is a placeholder implementation.")

    def view_draft(self):
        """Switch to draft view."""
        QMessageBox.information(self, "Draft View", "Draft view would be activated.\nThis is a placeholder implementation.")

    def view_zoom_dialog(self):
        """Show zoom dialog."""
        items = ["200%", "150%", "100%", "75%", "50%", "25%", "Page Width", "Whole Page", "Custom"]
        item, ok = QInputDialog.getItem(self, "Zoom", "Select zoom level:", items, 2, False)
        if ok and item:
            if item.endswith('%'):
                zoom_level = int(item.rstrip('%'))
                if self.current_editor():
                    self.current_editor().zoom_factor = zoom_level / 100.0
                    self.current_editor().update_zoom()
            elif item == "Page Width":
                self.view_page_width()
            elif item == "Whole Page":
                # Fit whole page in view
                if self.current_editor():
                    self.current_editor().zoom_factor = 0.85
                    self.current_editor().update_zoom()

    def view_page_width(self):
        """Zoom to fit page width."""
        if self.current_editor():
            # Calculate zoom to fit page width
            self.current_editor().zoom_factor = 1.0
            self.current_editor().update_zoom()
            QMessageBox.information(self, "Page Width", "Zoomed to fit page width.")

    def view_new_window(self):
        """Open a new window for the current document."""
        QMessageBox.information(self, "New Window", "A new window would open for the current document.\nThis is a placeholder implementation.")

    def view_split(self):
        """Split the editor view."""
        QMessageBox.information(self, "Split View", "The editor would be split into two panes.\nThis is a placeholder implementation.")

    def view_side_by_side(self):
        """View two documents side by side."""
        QMessageBox.information(self, "View Side by Side", "Two documents would be displayed side by side.\nThis is a placeholder implementation.")
