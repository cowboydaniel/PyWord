from PySide6.QtWidgets import QToolBar, QMenu
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon, QKeySequence

class MainToolBar(QToolBar):
    def __init__(self, parent=None):
        super().__init__("Main Toolbar", parent)
        self.setIconSize(QSize(16, 16))
        self.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.setup_actions()
    
    def setup_actions(self):
        # New document
        self.new_action = QAction("New", self)
        self.new_action.setIcon(self.style().standardIcon("SP_FileIcon"))
        self.new_action.setShortcut(QKeySequence.New)
        self.new_action.setToolTip("New (Ctrl+N)")
        
        # Open document
        self.open_action = QAction("Open", self)
        self.open_action.setIcon(self.style().standardIcon("SP_DialogOpenButton"))
        self.open_action.setShortcut(QKeySequence.Open)
        self.open_action.setToolTip("Open (Ctrl+O)")
        
        # Save document
        self.save_action = QAction("Save", self)
        self.save_action.setIcon(self.style().standardIcon("SP_DialogSaveButton"))
        self.save_action.setShortcut(QKeySequence.Save)
        self.save_action.setToolTip("Save (Ctrl+S)")
        
        # Save as document
        self.save_as_action = QAction("Save As", self)
        self.save_as_action.setIcon(self.style().standardIcon("SP_DialogSaveButton"))
        self.save_as_action.setShortcut(QKeySequence.SaveAs)
        self.save_as_action.setToolTip("Save As (Ctrl+Shift+S)")
        
        # Print
        self.print_action = QAction("Print", self)
        self.print_action.setIcon(self.style().standardIcon("SP_FileDialogDetailedView"))
        self.print_action.setShortcut(QKeySequence.Print)
        self.print_action.setToolTip("Print (Ctrl+P)")
        
        # Add actions to toolbar
        self.addAction(self.new_action)
        self.addAction(self.open_action)
        self.addAction(self.save_action)
        self.addAction(self.save_as_action)
        self.addSeparator()
        self.addAction(self.print_action)
        self.addSeparator()
        
        # Undo
        self.undo_action = QAction("Undo", self)
        self.undo_action.setIcon(self.style().standardIcon("SP_ArrowBack"))
        self.undo_action.setShortcut(QKeySequence.Undo)
        self.undo_action.setToolTip("Undo (Ctrl+Z)")
        
        # Redo
        self.redo_action = QAction("Redo", self)
        self.redo_action.setIcon(self.style().standardIcon("SP_ArrowForward"))
        self.redo_action.setShortcut(QKeySequence.Redo)
        self.redo_action.setToolTip("Redo (Ctrl+Y)")
        
        # Add undo/redo actions
        self.addAction(self.undo_action)
        self.addAction(self.redo_action)
        self.addSeparator()
        
        # Cut
        self.cut_action = QAction("Cut", self)
        self.cut_action.setIcon(self.style().standardIcon("SP_DialogCancelButton"))
        self.cut_action.setShortcut(QKeySequence.Cut)
        self.cut_action.setToolTip("Cut (Ctrl+X)")
        
        # Copy
        self.copy_action = QAction("Copy", self)
        self.copy_action.setIcon(self.style().standardIcon("SP_DialogNoButton"))
        self.copy_action.setShortcut(QKeySequence.Copy)
        self.copy_action.setToolTip("Copy (Ctrl+C)")
        
        # Paste
        self.paste_action = QAction("Paste", self)
        self.paste_action.setIcon(self.style().standardIcon("SP_DialogYesButton"))
        self.paste_action.setShortcut(QKeySequence.Paste)
        self.paste_action.setToolTip("Paste (Ctrl+V)")
        
        # Add clipboard actions
        self.addAction(self.cut_action)
        self.addAction(self.copy_action)
        self.addAction(self.paste_action)
        self.addSeparator()
        
        # Find
        self.find_action = QAction("Find", self)
        self.find_action.setIcon(self.style().standardIcon("SP_FileDialogContentsView"))
        self.find_action.setShortcut(QKeySequence.Find)
        self.find_action.setToolTip("Find (Ctrl+F)")
        
        # Replace
        self.replace_action = QAction("Replace", self)
        self.replace_action.setIcon(self.style().standardIcon("SP_BrowserReload"))
        self.replace_action.setShortcut(QKeySequence.Replace)
        self.replace_action.setToolTip("Replace (Ctrl+H)")
        
        # Add find/replace actions
        self.addAction(self.find_action)
        self.addAction(self.replace_action)
