from PySide6.QtWidgets import QToolBar, QComboBox
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon

class ReviewToolBar(QToolBar):
    def __init__(self, parent=None):
        super().__init__("Review Toolbar", parent)
        self.setIconSize(QSize(16, 16))
        self.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.setup_ui()
    
    def setup_ui(self):
        # Track Changes
        self.track_changes_action = QAction("Track Changes", self)
        self.track_changes_action.setCheckable(True)
        self.track_changes_action.setToolTip("Track Changes")
        
        # Show Markup
        self.show_markup_combo = QComboBox(self)
        self.show_markup_combo.addItems(["No Markup", "All Markup", "Simple Markup", "No Markup"])
        self.show_markup_combo.setToolTip("Show Markup")
        self.show_markup_combo.setMaximumWidth(120)
        
        # Previous Change
        self.previous_change_action = QAction("Previous Change", self)
        self.previous_change_action.setIcon(self.style().standardIcon("SP_ArrowUp"))
        self.previous_change_action.setToolTip("Previous Change")
        
        # Next Change
        self.next_change_action = QAction("Next Change", self)
        self.next_change_action.setIcon(self.style().standardIcon("SP_ArrowDown"))
        self.next_change_action.setToolTip("Next Change")
        
        # Accept Change
        self.accept_change_action = QAction("Accept", self)
        self.accept_change_action.setIcon(self.style().standardIcon("SP_DialogApplyButton"))
        self.accept_change_action.setToolTip("Accept Change")
        
        # Reject Change
        self.reject_change_action = QAction("Reject", self)
        self.reject_change_action.setIcon(self.style().standardIcon("SP_DialogCancelButton"))
        self.reject_change_action.setToolTip("Reject Change")
        
        # Add actions to toolbar
        self.addAction(self.track_changes_action)
        self.addWidget(self.show_markup_combo)
        self.addSeparator()
        self.addAction(self.previous_change_action)
        self.addAction(self.next_change_action)
        self.addSeparator()
        self.addAction(self.accept_change_action)
        self.addAction(self.reject_change_action)
        self.addSeparator()
        
        # New Comment
        self.new_comment_action = QAction("New Comment", self)
        self.new_comment_action.setIcon(self.style().standardIcon("SP_MessageBoxInformation"))
        self.new_comment_action.setToolTip("New Comment")
        
        # Delete Comment
        self.delete_comment_action = QAction("Delete Comment", self)
        self.delete_comment_action.setIcon(self.style().standardIcon("SP_TrashIcon"))
        self.delete_comment_action.setToolTip("Delete Comment")
        
        # Previous Comment
        self.previous_comment_action = QAction("Previous Comment", self)
        self.previous_comment_action.setIcon(self.style().standardIcon("SP_ArrowUp"))
        self.previous_comment_action.setToolTip("Previous Comment")
        
        # Next Comment
        self.next_comment_action = QAction("Next Comment", self)
        self.next_comment_action.setIcon(self.style().standardIcon("SP_ArrowDown"))
        self.next_comment_action.setToolTip("Next Comment")
        
        # Add comment actions
        self.addAction(self.new_comment_action)
        self.addAction(self.delete_comment_action)
        self.addSeparator()
        self.addAction(self.previous_comment_action)
        self.addAction(self.next_comment_action)
