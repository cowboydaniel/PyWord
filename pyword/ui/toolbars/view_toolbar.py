from PySide6.QtWidgets import QToolBar, QAction, QComboBox, QLabel
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon, QKeySequence

class ViewToolBar(QToolBar):
    def __init__(self, parent=None):
        super().__init__("View Toolbar", parent)
        self.setIconSize(QSize(16, 16))
        self.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.setup_ui()
    
    def setup_ui(self):
        # Zoom level
        self.zoom_label = QLabel("Zoom: ", self)
        self.addWidget(self.zoom_label)
        
        self.zoom_combo = QComboBox(self)
        self.zoom_combo.addItems(["50%", "75%", "100%", "150%", "200%", "Page Width", "Whole Page", "Many Pages"])
        self.zoom_combo.setCurrentText("100%")
        self.zoom_combo.setToolTip("Zoom")
        self.zoom_combo.setMaximumWidth(120)
        self.addWidget(self.zoom_combo)
        
        # Zoom In
        self.zoom_in_action = QAction("Zoom In", self)
        self.zoom_in_action.setIcon(self.style().standardIcon("SP_ArrowUp"))
        self.zoom_in_action.setShortcut(QKeySequence.ZoomIn)
        self.zoom_in_action.setToolTip("Zoom In (Ctrl++)")
        
        # Zoom Out
        self.zoom_out_action = QAction("Zoom Out", self)
        self.zoom_out_action.setIcon(self.style().standardIcon("SP_ArrowDown"))
        self.zoom_out_action.setShortcut(QKeySequence.ZoomOut)
        self.zoom_out_action.setToolTip("Zoom Out (Ctrl+-)")
        
        # Add zoom actions
        self.addAction(self.zoom_in_action)
        self.addAction(self.zoom_out_action)
        self.addSeparator()
        
        # View Modes
        self.view_mode_combo = QComboBox(self)
        self.view_mode_combo.addItems(["Print Layout", "Web Layout", "Outline", "Draft"])
        self.view_mode_combo.setToolTip("View Mode")
        self.view_mode_combo.setMaximumWidth(120)
        self.addWidget(self.view_mode_combo)
        self.addSeparator()
        
        # Ruler
        self.ruler_action = QAction("Ruler", self)
        self.ruler_action.setCheckable(True)
        self.ruler_action.setChecked(True)
        self.ruler_action.setToolTip("Ruler")
        
        # Gridlines
        self.gridlines_action = QAction("Gridlines", self)
        self.gridlines_action.setCheckable(True)
        self.gridlines_action.setToolTip("Gridlines")
        
        # Navigation Pane
        self.nav_pane_action = QAction("Navigation Pane", self)
        self.nav_pane_action.setCheckable(True)
        self.nav_pane_action.setShortcut("Ctrl+F1")
        self.nav_pane_action.setToolTip("Navigation Pane (Ctrl+F1)")
        
        # Add view actions
        self.addAction(self.ruler_action)
        self.addAction(self.gridlines_action)
        self.addAction(self.nav_pane_action)
        self.addSeparator()
        
        # Split Window
        self.split_action = QAction("Split", self)
        self.split_action.setCheckable(True)
        self.split_action.setToolTip("Split Window")
        
        # Side by Side
        self.side_by_side_action = QAction("Side by Side", self)
        self.side_by_side_action.setCheckable(True)
        self.side_by_side_action.setToolTip("View Side by Side")
        
        # Synchronous Scrolling
        self.sync_scrolling_action = QAction("Synchronous Scrolling", self)
        self.sync_scrolling_action.setCheckable(True)
        self.sync_scrolling_action.setToolTip("Synchronous Scrolling")
        
        # Reset Window Position
        self.reset_window_action = QAction("Reset Window Position", self)
        self.reset_window_action.setToolTip("Reset Window Position")
        
        # Add window management actions
        self.addAction(self.split_action)
        self.addAction(self.side_by_side_action)
        self.addAction(self.sync_scrolling_action)
        self.addAction(self.reset_window_action)
