from PySide6.QtWidgets import (QWidget, QVBoxLayout, QListWidget, QListWidgetItem, 
                             QPushButton, QMenu, QInputDialog, QLineEdit, QMessageBox)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QTextCharFormat, QTextBlockFormat, QFont, QColor, QAction

class StylesPanel(QWidget):
    """Panel for managing and applying text and paragraph styles."""
    
    # Signals
    style_applied = Signal(str)  # style_name
    style_modified = Signal(str, dict)  # style_name, properties
    style_created = Signal(str, dict)  # style_name, properties
    style_deleted = Signal(str)  # style_name
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.styles = {}
        self.setup_ui()
        self.load_default_styles()
    
    def setup_ui(self):
        """Initialize the styles panel UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        
        # Style filter
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Filter styles...")
        self.filter_edit.textChanged.connect(self.filter_styles)
        
        # Styles list
        self.styles_list = QListWidget()
        self.styles_list.setAlternatingRowColors(True)
        self.styles_list.itemDoubleClicked.connect(self.apply_style)
        self.styles_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.styles_list.customContextMenuRequested.connect(self.show_context_menu)
        
        # Buttons
        button_layout = QVBoxLayout()
        
        self.new_style_btn = QPushButton("New Style...")
        self.new_style_btn.clicked.connect(self.create_style)
        
        self.modify_style_btn = QPushButton("Modify Style...")
        self.modify_style_btn.clicked.connect(self.modify_style)
        
        self.delete_style_btn = QPushButton("Delete Style")
        self.delete_style_btn.clicked.connect(self.delete_style)
        
        # Add widgets to layout
        layout.addWidget(self.filter_edit)
        layout.addWidget(self.styles_list)
        button_layout.addWidget(self.new_style_btn)
        button_layout.addWidget(self.modify_style_btn)
        button_layout.addWidget(self.delete_style_btn)
        button_layout.addStretch()
        layout.addLayout(button_layout)
    
    def load_default_styles(self):
        """Load default styles."""
        # Clear existing styles
        self.styles.clear()
        self.styles_list.clear()
        
        # Add default styles
        self.add_style("Normal", {"font_family": "Arial", "font_size": 12, "bold": False, "italic": False})
        self.add_style("Heading 1", {"font_family": "Arial", "font_size": 16, "bold": True, "spacing_after": 12})
        self.add_style("Heading 2", {"font_family": "Arial", "font_size": 14, "bold": True, "italic": True, "spacing_after": 10})
        self.add_style("Title", {"font_family": "Arial", "font_size": 18, "bold": True, "alignment": "center"})
        self.add_style("Quote", {"font_family": "Times New Roman", "font_size": 12, "italic": True, "left_margin": 20, "right_margin": 20})
    
    def add_style(self, name, properties):
        """Add a style to the panel."""
        if name not in self.styles:
            self.styles[name] = properties
            item = QListWidgetItem(name)
            item.setData(Qt.UserRole, name)
            self.styles_list.addItem(item)
    
    def update_style(self, name, properties):
        """Update an existing style."""
        if name in self.styles:
            self.styles[name].update(properties)
            self.style_modified.emit(name, self.styles[name])
    
    def remove_style(self, name):
        """Remove a style from the panel."""
        if name in self.styles:
            del self.styles[name]
            for i in range(self.styles_list.count()):
                if self.styles_list.item(i).data(Qt.UserRole) == name:
                    self.styles_list.takeItem(i)
                    break
            self.style_deleted.emit(name)
    
    def get_style(self, name):
        """Get style properties by name."""
        return self.styles.get(name, {})
    
    def filter_styles(self, text):
        """Filter styles based on the filter text."""
        text = text.lower()
        for i in range(self.styles_list.count()):
            item = self.styles_list.item(i)
            item.setHidden(text not in item.text().lower())
    
    def apply_style(self, item):
        """Apply the selected style to the current selection."""
        style_name = item.data(Qt.UserRole)
        self.style_applied.emit(style_name)
    
    def create_style(self):
        """Create a new style."""
        name, ok = QInputDialog.getText(self, "New Style", "Enter style name:")
        if ok and name:
            if name in self.styles:
                QMessageBox.warning(self, "Error", f"A style named '{name}' already exists.")
                return
            
            # Default properties for new style
            properties = {
                "font_family": "Arial",
                "font_size": 12,
                "bold": False,
                "italic": False,
                "alignment": "left"
            }
            
            # TODO: Open style dialog to edit properties
            
            self.add_style(name, properties)
            self.style_created.emit(name, properties)
    
    def modify_style(self):
        """Modify the selected style."""
        selected = self.styles_list.currentItem()
        if not selected:
            return
            
        style_name = selected.data(Qt.UserRole)
        properties = self.styles[style_name].copy()
        
        # TODO: Open style dialog to edit properties
        # For now, just show a message
        QMessageBox.information(self, "Modify Style", f"Modifying style: {style_name}")
        
        # Update the style if modified
        if properties != self.styles[style_name]:
            self.update_style(style_name, properties)
    
    def delete_style(self):
        """Delete the selected style."""
        selected = self.styles_list.currentItem()
        if not selected:
            return
            
        style_name = selected.data(Qt.UserRole)
        
        # Don't allow deleting default styles
        if style_name in ["Normal", "Heading 1", "Heading 2"]:
            QMessageBox.warning(self, "Cannot Delete", "Default styles cannot be deleted.")
            return
        
        reply = QMessageBox.question(
            self, "Delete Style", 
            f"Are you sure you want to delete the style '{style_name}'?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.remove_style(style_name)
    
    def show_context_menu(self, position):
        """Show context menu for styles."""
        item = self.styles_list.itemAt(position)
        if not item:
            return
            
        menu = QMenu()
        
        apply_action = QAction("Apply", self)
        apply_action.triggered.connect(lambda: self.apply_style(item))
        
        modify_action = QAction("Modify...", self)
        modify_action.triggered.connect(self.modify_style)
        
        delete_action = QAction("Delete", self)
        delete_action.triggered.connect(self.delete_style)
        
        menu.addAction(apply_action)
        menu.addSeparator()
        menu.addAction(modify_action)
        menu.addAction(delete_action)
        
        menu.exec_(self.styles_list.mapToGlobal(position))
