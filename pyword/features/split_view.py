from PySide6.QtCore import Qt, Signal, QPoint, QObject
from PySide6.QtWidgets import (
    QSplitter, QTextEdit, QWidget, QVBoxLayout, QHBoxLayout,
    QToolBar, QAction, QLabel, QScrollBar, QApplication
)
from PySide6.QtGui import QIcon

class SplitViewManager(QObject):
    """Manages split view functionality for the document editor."""
    viewCountChanged = Signal(int)  # Emitted when number of views changes
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.views = []
        self.current_view_index = 0
        self.synchronize_views = True
        self.split_orientation = Qt.Orientation.Vertical
    
    def create_split_view(self, editor_widget_factory):
        """Create a new split view container."""
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Create the splitter
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setHandleWidth(2)
        splitter.setChildrenCollapsible(False)
        
        # Create initial editor view
        editor = editor_widget_factory()
        self.views.append(editor)
        splitter.addWidget(editor)
        
        # Store the splitter and editor factory for later use
        container.splitter = splitter
        container.editor_factory = editor_widget_factory
        
        layout.addWidget(splitter)
        self.viewCountChanged.emit(len(self.views))
        
        return container
    
    def split_view(self, container, orientation=None):
        """Split the current view."""
        if orientation is None:
            orientation = self.split_orientation
        
        # Create a new editor
        new_editor = container.editor_factory()
        self.views.append(new_editor)
        
        # Add to splitter
        container.splitter.addWidget(new_editor)
        
        # Set orientation if different
        if container.splitter.orientation() != orientation:
            container.splitter.setOrientation(orientation)
        
        # Make sure the split is even
        sizes = container.splitter.sizes()
        if sizes:
            new_size = sum(sizes) // (len(sizes) + 1)
            new_sizes = [new_size] * (len(sizes) + 1)
            container.splitter.setSizes(new_sizes)
        
        self.viewCountChanged.emit(len(self.views))
        return new_editor
    
    def close_view(self, container, editor):
        """Close a view from the splitter."""
        if len(self.views) <= 1:
            return False  # Can't close the last view
        
        if editor in self.views:
            self.views.remove(editor)
            editor.setParent(None)
            editor.deleteLater()
            self.viewCountChanged.emit(len(self.views))
            return True
        return False
    
    def close_all_views(self, container):
        """Close all views except the first one."""
        if len(self.views) <= 1:
            return
        
        # Keep the first view
        first_view = self.views[0]
        
        # Remove all other views
        for view in self.views[1:]:
            view.setParent(None)
            view.deleteLater()
        
        # Reset views list
        self.views = [first_view]
        
        # Reset splitter
        container.splitter = QSplitter(Qt.Orientation.Vertical)
        container.splitter.addWidget(first_view)
        
        # Replace the layout's widget with the new splitter
        old_splitter = container.layout().itemAt(0).widget()
        container.layout().replaceWidget(old_splitter, container.splitter)
        old_splitter.deleteLater()
        
        self.viewCountChanged.emit(1)
    
    def set_synchronization(self, enabled):
        """Enable or disable view synchronization."""
        self.synchronize_views = enabled
        
        if enabled and len(self.views) > 1:
            # Reconnect scroll bars for synchronization
            self._connect_scroll_bars()
    
    def _connect_scroll_bars(self):
        """Connect scroll bars of all views for synchronization."""
        if not self.views:
            return
        
        # Disconnect any existing connections
        for view in self.views:
            try:
                view.verticalScrollBar().valueChanged.disconnect()
                view.horizontalScrollBar().valueChanged.disconnect()
            except:
                pass
        
        if not self.synchronize_views or len(self.views) <= 1:
            return
        
        # Connect all vertical scroll bars
        for i, view in enumerate(self.views):
            if i == 0:
                # First view is the master
                view.verticalScrollBar().valueChanged.connect(
                    lambda v, vbar=view.verticalScrollBar(): 
                        self._sync_scroll_bars(vbar, 'vertical')
                )
                view.horizontalScrollBar().valueChanged.connect(
                    lambda v, hbar=view.horizontalScrollBar(): 
                        self._sync_scroll_bars(hbar, 'horizontal')
                )
            else:
                # Other views follow the first one
                self.views[0].verticalScrollBar().valueChanged.connect(
                    view.verticalScrollBar().setValue
                )
                self.views[0].horizontalScrollBar().valueChanged.connect(
                    view.horizontalScrollBar().setValue
                )
    
    def _sync_scroll_bars(self, source_bar, orientation):
        """Synchronize scroll bars when in synchronized mode."""
        if not self.synchronize_views or len(self.views) <= 1:
            return
        
        # Block signals to prevent recursion
        for view in self.views[1:]:  # Skip the first view (source)
            if orientation == 'vertical':
                view.verticalScrollBar().blockSignals(True)
                view.verticalScrollBar().setValue(source_bar.value())
                view.verticalScrollBar().blockSignals(False)
            else:  # horizontal
                view.horizontalScrollBar().blockSignals(True)
                view.horizontalScrollBar().setValue(source_bar.value())
                view.horizontalScrollBar().blockSignals(False)
    
    def set_current_view(self, view):
        """Set the current active view."""
        if view in self.views:
            self.current_view_index = self.views.index(view)
    
    def get_current_view(self):
        """Get the current active view."""
        if 0 <= self.current_view_index < len(self.views):
            return self.views[self.current_view_index]
        return None if not self.views else self.views[0]
    
    def get_views(self):
        """Get all views."""
        return self.views
    
    def count(self):
        """Get the number of views."""
        return len(self.views)
    
    def set_split_orientation(self, orientation):
        """Set the split orientation (Qt.Horizontal or Qt.Vertical)."""
        if orientation in (Qt.Horizontal, Qt.Vertical):
            self.split_orientation = orientation
            return True
        return False


class SplitViewToolBar(QToolBar):
    """Toolbar for split view controls."""
    def __init__(self, split_view_manager, parent=None):
        super().__init__("Split View", parent)
        self.split_view_manager = split_view_manager
        self.setMovable(False)
        self.setIconSize(QSize(16, 16))
        
        self._create_actions()
        self._setup_ui()
        
        # Connect signals
        self.split_view_manager.viewCountChanged.connect(self._on_view_count_changed)
    
    def _create_actions(self):
        """Create actions for the toolbar."""
        # Split actions
        self.split_horizontal_action = QAction(
            QIcon.fromTheme("view-split-left-right"),
            "Split Horizontally",
            self,
            triggered=lambda: self.split_view(Qt.Orientation.Horizontal)
        )
        
        self.split_vertical_action = QAction(
            QIcon.fromTheme("view-split-top-bottom"),
            "Split Vertically",
            self,
            triggered=lambda: self.split_view(Qt.Orientation.Vertical)
        )
        
        # Close action
        self.close_view_action = QAction(
            QIcon.fromTheme("window-close"),
            "Close Current View",
            self,
            triggered=self.close_current_view
        )
        
        # Close all but this
        self.close_others_action = QAction(
            QIcon.fromTheme("view-close"),
            "Close Other Views",
            self,
            triggered=self.close_other_views
        )
        
        # Synchronize scrolling
        self.sync_action = QAction(
            QIcon.fromTheme("view-synchronize"),
            "Synchronize Scrolling",
            self,
            checkable=True,
            checked=True,
            toggled=self.split_view_manager.set_synchronization
        )
        
        # View mode actions
        self.view_mode_group = QActionGroup(self)
        
        self.print_layout_action = QAction(
            "Print Layout",
            self.view_mode_group,
            checkable=True,
            checked=True
        )
        
        self.web_layout_action = QAction(
            "Web Layout",
            self.view_mode_group,
            checkable=True
        )
        
        self.draft_view_action = QAction(
            "Draft View",
            self.view_mode_group,
            checkable=True
        )
        
        # Connect view mode actions
        self.print_layout_action.triggered.connect(
            lambda: self.set_view_mode('print'))
        self.web_layout_action.triggered.connect(
            lambda: self.set_view_mode('web'))
        self.draft_view_action.triggered.connect(
            lambda: self.set_view_mode('draft'))
    
    def _setup_ui(self):
        """Set up the toolbar UI."""
        # Add split actions
        self.addAction(self.split_horizontal_action)
        self.addAction(self.split_vertical_action)
        self.addSeparator()
        
        # Add close actions
        self.addAction(self.close_view_action)
        self.addAction(self.close_others_action)
        self.addSeparator()
        
        # Add sync action
        self.addAction(self.sync_action)
        self.addSeparator()
        
        # Add view mode actions
        self.addWidget(QLabel("View: "))
        self.addAction(self.print_layout_action)
        self.addAction(self.web_layout_action)
        self.addAction(self.draft_view_action)
        
        # Update action states
        self._on_view_count_changed(self.split_view_manager.count())
    
    def split_view(self, orientation):
        """Split the current view."""
        current_view = self.split_view_manager.get_current_view()
        if not current_view:
            return
            
        container = current_view.parent().parent()  # Get the container widget
        if container and hasattr(container, 'splitter'):
            self.split_view_manager.set_split_orientation(orientation)
            new_view = self.split_view_manager.split_view(container, orientation)
            
            # Copy document and cursor position from current view
            if hasattr(current_view, 'document') and hasattr(new_view, 'document'):
                new_view.document().setPlainText(current_view.document().toPlainText())
                new_view.setFocus()
    
    def close_current_view(self):
        """Close the current view."""
        current_view = self.split_view_manager.get_current_view()
        if current_view and self.split_view_manager.count() > 1:
            container = current_view.parent().parent()
            if container and hasattr(container, 'splitter'):
                self.split_view_manager.close_view(container, current_view)
    
    def close_other_views(self):
        """Close all views except the current one."""
        current_view = self.split_view_manager.get_current_view()
        if not current_view or self.split_view_manager.count() <= 1:
            return
            
        container = current_view.parent().parent()
        if container and hasattr(container, 'splitter'):
            # Keep only the current view
            for view in self.split_view_manager.get_views()[:]:
                if view != current_view:
                    self.split_view_manager.close_view(container, view)
    
    def set_view_mode(self, mode):
        """Set the view mode for all views."""
        for view in self.split_view_manager.get_views():
            if hasattr(view, 'set_view_mode'):
                view.set_view_mode(mode)
    
    def _on_view_count_changed(self, count):
        """Update UI when the number of views changes."""
        has_multiple_views = count > 1
        self.close_view_action.setEnabled(has_multiple_views)
        self.close_others_action.setEnabled(has_multiple_views)
        self.sync_action.setEnabled(has_multiple_views)


class ViewModeMixin:
    """Mixin class for view mode functionality."""
    def __init__(self):
        self._view_mode = 'print'  # 'print', 'web', or 'draft'
        self._setup_view_mode()
    
    def set_view_mode(self, mode):
        """Set the view mode."""
        if mode not in ('print', 'web', 'draft'):
            return
            
        self._view_mode = mode
        self._setup_view_mode()
    
    def _setup_view_mode(self):
        """Configure the view based on the current mode."""
        if not hasattr(self, 'document'):
            return
            
        document = self.document()
        
        if self._view_mode == 'print':
            # Print layout - show page boundaries, margins, etc.
            document.setDocumentMargin(72)  # 1 inch margins
            self.setViewportMargins(40, 40, 40, 40)  # Space for page shadow
            self.setFrameShape(QFrame.NoFrame)
            
            # Set up page background
            palette = self.palette()
            palette.setColor(self.backgroundRole(), Qt.gray)
            self.setPalette(palette)
            
            # Enable page mode
            document.setPageSize(self.viewport().size() - QSize(80, 80))
            
        elif self._view_mode == 'web':
            # Web layout - continuous scrolling, no page breaks
            document.setDocumentMargin(40)  # Reasonable margins
            self.setViewportMargins(0, 0, 0, 0)
            self.setFrameShape(QFrame.NoFrame)
            
            # Reset background
            palette = self.palette()
            palette.setColor(self.backgroundRole(), Qt.white)
            self.setPalette(palette)
            
            # Use viewport width
            document.setTextWidth(self.viewport().width() - 80)
            
        elif self._view_mode == 'draft':
            # Draft view - focus on content, minimal UI
            document.setDocumentMargin(20)
            self.setViewportMargins(0, 0, 0, 0)
            self.setFrameShape(QFrame.NoFrame)
            
            # Monospace font for draft view
            font = QFont("Courier", 10)
            document.setDefaultFont(font)
            
            # Light background
            palette = self.palette()
            palette.setColor(self.backgroundRole(), QColor(240, 240, 240))
            self.setPalette(palette)
            
            # Full width text
            document.setTextWidth(self.viewport().width() - 40)
    
    def resizeEvent(self, event):
        """Handle resize events to update the view mode layout."""
        super().resizeEvent(event)
        self._setup_view_mode()


# Example usage
if __name__ == "__main__":
    import sys
    from PySide6.QtWidgets import QApplication, QMainWindow, QTextEdit
    
    class Editor(QTextEdit, ViewModeMixin):
        def __init__(self, parent=None):
            super().__init__(parent)
            ViewModeMixin.__init__(self)
    
    app = QApplication(sys.argv)
    
    # Create main window
    window = QMainWindow()
    window.setWindowTitle("Split View Example")
    
    # Create split view manager
    split_view_manager = SplitViewManager()
    
    # Create editor factory function
    def create_editor():
        editor = Editor()
        editor.setPlainText("This is an editor view.\n" * 100)
        return editor
    
    # Create split view container
    container = split_view_manager.create_split_view(create_editor)
    
    # Create toolbar
    toolbar = SplitViewToolBar(split_view_manager)
    
    # Set up main layout
    main_widget = QWidget()
    layout = QVBoxLayout(main_widget)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(0)
    layout.addWidget(toolbar)
    layout.addWidget(container)
    
    window.setCentralWidget(main_widget)
    window.resize(800, 600)
    window.show()
    
    sys.exit(app.exec())
