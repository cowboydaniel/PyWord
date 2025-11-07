from PySide6.QtWidgets import QToolBar, QMenu, QComboBox
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon, QKeySequence
from ...features.headers_footers import HeaderFooterType

class HeaderFooterToolBar(QToolBar):
    """Toolbar for managing headers and footers in the document."""
    
    def __init__(self, parent=None):
        super().__init__("Header & Footer", parent)
        self.setIconSize(QSize(16, 16))
        self.setToolButtonStyle(Qt.ToolButtonIconOnly)
        
        # Initialize the current document
        self._current_document = None
        
        # Setup actions and connections
        self.setup_actions()
    
    def set_current_document(self, document):
        """Set the current active document."""
        self._current_document = document
        self.update_actions()
    
    def update_actions(self):
        """Update the state of toolbar actions based on current document."""
        has_document = self._current_document is not None
        
        # Enable/disable actions based on document state
        for action in self.actions():
            if action not in [self.header_menu.menuAction(), self.footer_menu.menuAction()]:
                action.setEnabled(has_document)
    
    def setup_actions(self):
        """Initialize all header/footer related actions."""
        # Header menu
        self.header_menu = QMenu("Header", self)
        
        # Add header actions
        self.header_insert = QAction("Insert Header", self)
        self.header_edit = QAction("Edit Header", self)
        self.header_remove = QAction("Remove Header", self)
        self.header_first_page = QAction("Different First Page", self, checkable=True)
        self.header_odd_even = QAction("Different Odd & Even Pages", self, checkable=True)
        
        # Add header actions to menu
        self.header_menu.addAction(self.header_insert)
        self.header_menu.addAction(self.header_edit)
        self.header_menu.addAction(self.header_remove)
        self.header_menu.addSeparator()
        self.header_menu.addAction(self.header_first_page)
        self.header_menu.addAction(self.header_odd_even)
        
        # Footer menu
        self.footer_menu = QMenu("Footer", self)
        
        # Add footer actions
        self.footer_insert = QAction("Insert Footer", self)
        self.footer_edit = QAction("Edit Footer", self)
        self.footer_remove = QAction("Remove Footer", self)
        self.footer_first_page = QAction("Different First Page", self, checkable=True)
        self.footer_odd_even = QAction("Different Odd & Even Pages", self, checkable=True)
        
        # Add footer actions to menu
        self.footer_menu.addAction(self.footer_insert)
        self.footer_menu.addAction(self.footer_edit)
        self.footer_menu.addAction(self.footer_remove)
        self.footer_menu.addSeparator()
        self.footer_menu.addAction(self.footer_first_page)
        self.footer_menu.addAction(self.footer_odd_even)
        
        # Page number dropdown
        self.page_number_menu = QMenu("Page Number", self)
        self.page_number_top = QAction("Top of Page", self)
        self.page_number_bottom = QAction("Bottom of Page", self)
        self.page_number_margins = QAction("Page Margins", self)
        self.page_number_current = QAction("Current Position", self)
        self.page_number_format = QAction("Format Page Numbers...", self)
        self.page_number_remove = QAction("Remove Page Numbers", self)
        
        # Add page number actions to menu
        self.page_number_menu.addAction(self.page_number_top)
        self.page_number_menu.addAction(self.page_number_bottom)
        self.page_number_menu.addAction(self.page_number_margins)
        self.page_number_menu.addAction(self.page_number_current)
        self.page_number_menu.addSeparator()
        self.page_number_menu.addAction(self.page_number_format)
        self.page_number_menu.addAction(self.page_number_remove)
        
        # Add widgets to toolbar
        self.addAction(self.header_menu.menuAction())
        self.addAction(self.footer_menu.menuAction())
        self.addSeparator()
        self.addAction(self.page_number_menu.menuAction())
        
        # Connect signals
        self.setup_connections()
        
        # Disable all actions by default (until a document is loaded)
        self.update_actions()
    
    def setup_connections(self):
        """Connect all action signals to their respective slots."""
        # Header actions
        self.header_insert.triggered.connect(self.on_header_insert)
        self.header_edit.triggered.connect(self.on_header_edit)
        self.header_remove.triggered.connect(self.on_header_remove)
        self.header_first_page.toggled.connect(self.on_header_first_page_toggled)
        self.header_odd_even.toggled.connect(self.on_header_odd_even_toggled)
        
        # Footer actions
        self.footer_insert.triggered.connect(self.on_footer_insert)
        self.footer_edit.triggered.connect(self.on_footer_edit)
        self.footer_remove.triggered.connect(self.on_footer_remove)
        self.footer_first_page.toggled.connect(self.on_footer_first_page_toggled)
        self.footer_odd_even.toggled.connect(self.on_footer_odd_even_toggled)
        
        # Page number actions
        self.page_number_top.triggered.connect(lambda: self.on_page_number_insert("top"))
        self.page_number_bottom.triggered.connect(lambda: self.on_page_number_insert("bottom"))
        self.page_number_margins.triggered.connect(lambda: self.on_page_number_insert("margins"))
        self.page_number_current.triggered.connect(lambda: self.on_page_number_insert("current"))
        self.page_number_format.triggered.connect(self.on_page_number_format)
        self.page_number_remove.triggered.connect(self.on_page_number_remove)
    
    # Header/Footer action handlers
    def on_header_insert(self):
        if self._current_document:
            self._current_document.header_footer_manager.show_header_dialog(HeaderFooterType.HEADER)
    
    def on_header_edit(self):
        if self._current_document:
            self._current_document.header_footer_manager.edit_header_footer(HeaderFooterType.HEADER)
    
    def on_header_remove(self):
        if self._current_document:
            self._current_document.header_footer_manager.remove_header_footer(HeaderFooterType.HEADER)
    
    def on_header_first_page_toggled(self, checked):
        if self._current_document:
            self._current_document.header_footer_manager.set_different_first_page(checked)
    
    def on_header_odd_even_toggled(self, checked):
        if self._current_document:
            self._current_document.header_footer_manager.set_different_odd_even_pages(checked)
    
    def on_footer_insert(self):
        if self._current_document:
            self._current_document.header_footer_manager.show_header_dialog(HeaderFooterType.FOOTER)
    
    def on_footer_edit(self):
        if self._current_document:
            self._current_document.header_footer_manager.edit_header_footer(HeaderFooterType.FOOTER)
    
    def on_footer_remove(self):
        if self._current_document:
            self._current_document.header_footer_manager.remove_header_footer(HeaderFooterType.FOOTER)
    
    def on_footer_first_page_toggled(self, checked):
        # Same as header since it's a document-wide setting
        self.on_header_first_page_toggled(checked)
    
    def on_footer_odd_even_toggled(self, checked):
        # Same as header since it's a document-wide setting
        self.on_header_odd_even_toggled(checked)
    
    # Page number handlers
    def on_page_number_insert(self, position):
        if not self._current_document:
            return
            
        manager = self._current_document.header_footer_manager
        if position in ["top", "bottom"]:
            manager.insert_page_number(position)
        elif position == "margins":
            manager.insert_page_number_margins()
        else:  # current position
            manager.insert_page_number_at_cursor()
    
    def on_page_number_format(self):
        if self._current_document:
            self._current_document.header_footer_manager.format_page_numbers()
    
    def on_page_number_remove(self):
        if self._current_document:
            self._current_document.header_footer_manager.remove_page_numbers()
