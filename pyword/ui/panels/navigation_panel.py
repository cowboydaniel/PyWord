from PySide6.QtWidgets import (QWidget, QVBoxLayout, QTreeWidget, QTreeWidgetItem, 
                             QTabWidget, QListWidget, QLineEdit, QLabel, QHBoxLayout)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QIcon

class NavigationPanel(QWidget):
    """Navigation panel with document outline, pages, and search results."""
    
    # Signals
    navigation_requested = Signal(int, int)  # line, position
    search_requested = Signal(str, bool, bool)  # text, case_sensitive, whole_words
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
    
    def setup_ui(self):
        """Initialize the navigation panel UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Search box
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Search document...")
        self.search_edit.returnPressed.connect(self.on_search)
        
        # Tab widget for different navigation views
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.West)
        
        # Headings tab
        self.headings_tree = QTreeWidget()
        self.headings_tree.setHeaderHidden(True)
        self.headings_tree.itemClicked.connect(self.on_heading_clicked)
        self.tab_widget.addTab(self.headings_tree, "Headings")
        
        # Pages tab
        self.pages_list = QListWidget()
        self.pages_list.itemClicked.connect(self.on_page_clicked)
        self.tab_widget.addTab(self.pages_list, "Pages")
        
        # Search results tab
        self.search_results = QListWidget()
        self.search_results.itemClicked.connect(self.on_search_result_clicked)
        self.tab_widget.addTab(self.search_results, "Results")
        
        # Add widgets to layout
        layout.addWidget(self.search_edit)
        layout.addWidget(self.tab_widget)
    
    def update_headings(self, headings):
        """Update the headings tree with document headings.
        
        Args:
            headings: List of tuples (level, text, position)
        """
        self.headings_tree.clear()
        
        # Create a dictionary to hold parent items for each level
        parents = {0: None}
        
        for level, text, position in headings:
            item = QTreeWidgetItem([text])
            item.setData(0, Qt.UserRole, position)
            
            # Find the parent for this level
            parent_level = level - 1
            while parent_level >= 0 and parent_level not in parents:
                parent_level -= 1
            
            parent = parents.get(parent_level)
            if parent is not None:
                parent.addChild(item)
            else:
                self.headings_tree.addTopLevelItem(item)
            
            # Update parents dictionary
            parents[level] = item
        
        # Expand all items by default
        self.headings_tree.expandAll()
    
    def update_pages(self, pages):
        """Update the pages list with document pages.
        
        Args:
            pages: List of tuples (page_number, preview_text)
        """
        self.pages_list.clear()
        for page_num, preview_text in pages:
            self.pages_list.addItem(f"Page {page_num}: {preview_text[:50]}...")
    
    def update_search_results(self, results):
        """Update the search results list.
        
        Args:
            results: List of tuples (line_number, preview_text, position)
        """
        self.search_results.clear()
        for line_num, preview_text, position in results:
            item = QListWidgetItem(f"Line {line_num}: {preview_text}")
            item.setData(Qt.UserRole, (line_num, position))
            self.search_results.addItem(item)
    
    def on_heading_clicked(self, item, column):
        """Handle heading click event."""
        position = item.data(0, Qt.UserRole)
        if position is not None:
            self.navigation_requested.emit(-1, position)
    
    def on_page_clicked(self, item):
        """Handle page click event."""
        # TODO: Implement page navigation
        pass
    
    def on_search_result_clicked(self, item):
        """Handle search result click event."""
        data = item.data(Qt.UserRole)
        if data:
            line_num, position = data
            self.navigation_requested.emit(line_num, position)
    
    def on_search(self):
        """Handle search action."""
        text = self.search_edit.text().strip()
        if text:
            # Emit search signal with current search options
            case_sensitive = False  # Get from UI
            whole_words = False     # Get from UI
            self.search_requested.emit(text, case_sensitive, whole_words)
            
            # Switch to search results tab
            self.tab_widget.setCurrentWidget(self.search_results)
