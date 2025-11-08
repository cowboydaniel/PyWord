"""Ribbon interface for PyWord - modern Microsoft Office-style UI."""

from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QToolButton,
                               QLabel, QFrame, QScrollArea, QSizePolicy, QPushButton,
                               QButtonGroup, QGridLayout, QSpacerItem, QMenu, QStackedWidget)
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtGui import QIcon, QAction, QFont, QColor


class RibbonTab(QWidget):
    """A single tab in the ribbon interface."""

    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.title = title
        self.groups = []
        self.setup_ui()

    def setup_ui(self):
        """Initialize the tab UI."""
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(5, 5, 5, 5)
        self.layout.setSpacing(10)
        self.layout.addStretch()

    def add_group(self, group: 'RibbonGroup'):
        """Add a group to this tab."""
        self.groups.append(group)
        self.layout.insertWidget(len(self.groups) - 1, group)

        # Add separator after group (except for last group)
        separator = QFrame()
        separator.setFrameShape(QFrame.VLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setFixedWidth(1)
        self.layout.insertWidget(len(self.groups) * 2 - 1, separator)


class RibbonGroup(QWidget):
    """A group of related controls in a ribbon tab."""

    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.title = title
        self.buttons = []
        self.setup_ui()

    def setup_ui(self):
        """Initialize the group UI."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 3, 8, 3)
        main_layout.setSpacing(2)

        # Content area
        self.content_layout = QGridLayout()
        self.content_layout.setSpacing(4)
        main_layout.addLayout(self.content_layout)

        # Group title
        title_label = QLabel(self.title)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("""
            font-size: 11px;
            color: #605E5C;
            font-weight: normal;
            padding-top: 2px;
        """)
        main_layout.addWidget(title_label)

        self.current_row = 0
        self.current_col = 0

    def add_large_button(self, icon: QIcon, text: str, tooltip: str = "") -> QToolButton:
        """Add a large button (with text below icon)."""
        button = QToolButton()
        button.setIcon(icon)
        button.setText(text)
        button.setToolTip(tooltip or text)
        button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        button.setIconSize(QSize(32, 32))
        button.setFixedSize(68, 68)
        button.setStyleSheet("""
            QToolButton {
                border: 1px solid transparent;
                border-radius: 2px;
                padding: 4px;
                background: transparent;
                color: #323130;
                font-size: 11px;
            }
            QToolButton:hover {
                background-color: #F3F2F1;
                border: 1px solid #EDEBE9;
            }
            QToolButton:pressed {
                background-color: #EDEBE9;
                border: 1px solid #E1DFDD;
            }
        """)

        self.content_layout.addWidget(button, 0, self.current_col, 3, 1)
        self.current_col += 1
        self.buttons.append(button)
        return button

    def add_small_button(self, icon: QIcon, text: str, tooltip: str = "") -> QToolButton:
        """Add a small button (with text beside icon)."""
        button = QToolButton()
        button.setIcon(icon)
        button.setText(text)
        button.setToolTip(tooltip or text)
        button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        button.setIconSize(QSize(16, 16))
        button.setFixedHeight(24)
        button.setMinimumWidth(60)
        button.setStyleSheet("""
            QToolButton {
                border: 1px solid transparent;
                border-radius: 2px;
                padding: 3px 8px;
                background: transparent;
                color: #323130;
                font-size: 12px;
                text-align: left;
            }
            QToolButton:hover {
                background-color: #F3F2F1;
                border: 1px solid #EDEBE9;
            }
            QToolButton:pressed {
                background-color: #EDEBE9;
                border: 1px solid #E1DFDD;
            }
        """)

        self.content_layout.addWidget(button, self.current_row, self.current_col)
        self.current_row += 1

        if self.current_row >= 3:
            self.current_row = 0
            self.current_col += 1

        self.buttons.append(button)
        return button

    def add_widget(self, widget: QWidget, row_span: int = 1, col_span: int = 1):
        """Add a custom widget to the group."""
        self.content_layout.addWidget(widget, self.current_row, self.current_col,
                                     row_span, col_span)
        self.current_row += row_span

        if self.current_row >= 3:
            self.current_row = 0
            self.current_col += col_span


class RibbonBar(QWidget):
    """Modern ribbon interface bar."""

    tab_changed = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.tabs = []
        self.tab_buttons = []
        self.current_tab_index = 0
        self.setup_ui()

    def setup_ui(self):
        """Initialize the ribbon UI."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Tab bar
        self.tab_bar_widget = QWidget()
        self.tab_bar_widget.setStyleSheet("""
            QWidget {
                background-color: #FFFFFF;
                border-bottom: 1px solid #D2D0CE;
            }
        """)
        self.tab_bar_layout = QHBoxLayout(self.tab_bar_widget)
        self.tab_bar_layout.setContentsMargins(8, 0, 8, 0)
        self.tab_bar_layout.setSpacing(2)
        self.tab_bar_layout.addStretch()
        main_layout.addWidget(self.tab_bar_widget)

        # Content area with stacked widget for tabs
        self.content_widget = QWidget()
        self.content_widget.setStyleSheet("""
            QWidget {
                background-color: #FFFFFF;
                border-bottom: 1px solid #D2D0CE;
            }
        """)
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(0)
        main_layout.addWidget(self.content_widget)

        # Stacked widget for tabs wrapped in scroll area
        self.stacked_widget = QStackedWidget()

        # Scroll area for tabs
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll_area.setWidget(self.stacked_widget)
        self.scroll_area.setMaximumHeight(120)  # Limit ribbon height
        self.content_layout.addWidget(self.scroll_area)

    def add_tab(self, tab: RibbonTab) -> int:
        """Add a tab to the ribbon."""
        self.tabs.append(tab)

        # Add tab to stacked widget
        self.stacked_widget.addWidget(tab)

        # Create tab button
        button = QPushButton(tab.title)
        button.setCheckable(True)
        button.setFlat(True)
        button.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                border: none;
                border-bottom: 2px solid transparent;
                background: transparent;
                font-weight: normal;
                font-size: 13px;
                color: #323130;
            }
            QPushButton:hover {
                background-color: #F3F2F1;
            }
            QPushButton:checked {
                border-bottom: 2px solid #0078D4;
                background-color: #FFFFFF;
                color: #0078D4;
            }
        """)

        tab_index = len(self.tabs) - 1
        button.clicked.connect(lambda: self.set_current_tab(tab_index))

        self.tab_buttons.append(button)
        self.tab_bar_layout.insertWidget(len(self.tab_buttons) - 1, button)

        # Set first tab as active
        if len(self.tabs) == 1:
            button.setChecked(True)
            self.stacked_widget.setCurrentIndex(0)

        return tab_index

    def set_current_tab(self, index: int):
        """Set the current active tab."""
        if 0 <= index < len(self.tabs):
            # Uncheck all buttons
            for btn in self.tab_buttons:
                btn.setChecked(False)

            # Check current button
            self.tab_buttons[index].setChecked(True)

            # Show tab content using stacked widget
            self.stacked_widget.setCurrentIndex(index)
            self.current_tab_index = index
            self.tab_changed.emit(index)

    def create_file_tab(self) -> RibbonTab:
        """Create the File tab with document operations (backstage view)."""
        tab = RibbonTab("File")

        # Document operations group
        file_group = RibbonGroup("Document")
        file_group.add_large_button(QIcon(), "New", "Create new document")
        file_group.add_large_button(QIcon(), "Open", "Open document")
        file_group.add_large_button(QIcon(), "Save", "Save document")
        file_group.add_large_button(QIcon(), "Save As", "Save document as")
        tab.add_group(file_group)

        # Print group
        print_group = RibbonGroup("Print")
        print_group.add_large_button(QIcon(), "Print", "Print document")
        print_group.add_large_button(QIcon(), "Preview", "Print preview")
        tab.add_group(print_group)

        # Share group
        share_group = RibbonGroup("Share")
        share_group.add_large_button(QIcon(), "Export", "Export document")
        share_group.add_small_button(QIcon(), "PDF", "Export as PDF")
        tab.add_group(share_group)

        return tab

    def create_home_tab(self) -> RibbonTab:
        """Create the Home tab with common formatting options."""
        tab = RibbonTab("Home")

        # Clipboard group
        clipboard_group = RibbonGroup("Clipboard")
        # Note: Icons would be loaded from resources in a real implementation
        clipboard_group.add_large_button(QIcon(), "Paste", "Paste from clipboard")
        clipboard_group.add_small_button(QIcon(), "Cut", "Cut to clipboard")
        clipboard_group.add_small_button(QIcon(), "Copy", "Copy to clipboard")
        clipboard_group.add_small_button(QIcon(), "Format", "Format painter")
        tab.add_group(clipboard_group)

        # Font group
        font_group = RibbonGroup("Font")
        font_group.add_small_button(QIcon(), "Bold", "Bold")
        font_group.add_small_button(QIcon(), "Italic", "Italic")
        font_group.add_small_button(QIcon(), "Underline", "Underline")
        font_group.add_small_button(QIcon(), "Strike", "Strikethrough")
        font_group.add_small_button(QIcon(), "Color", "Font color")
        font_group.add_small_button(QIcon(), "Highlight", "Highlight")
        tab.add_group(font_group)

        # Paragraph group
        paragraph_group = RibbonGroup("Paragraph")
        paragraph_group.add_small_button(QIcon(), "Bullets", "Bullets")
        paragraph_group.add_small_button(QIcon(), "Numbering", "Numbering")
        paragraph_group.add_small_button(QIcon(), "Align Left", "Align left")
        paragraph_group.add_small_button(QIcon(), "Center", "Center")
        paragraph_group.add_small_button(QIcon(), "Align Right", "Align right")
        paragraph_group.add_small_button(QIcon(), "Justify", "Justify")
        tab.add_group(paragraph_group)

        # Styles group
        styles_group = RibbonGroup("Styles")
        styles_group.add_small_button(QIcon(), "Heading 1", "Heading 1")
        styles_group.add_small_button(QIcon(), "Heading 2", "Heading 2")
        styles_group.add_small_button(QIcon(), "Normal", "Normal text")
        tab.add_group(styles_group)

        return tab

    def create_insert_tab(self) -> RibbonTab:
        """Create the Insert tab."""
        tab = RibbonTab("Insert")

        # Tables group
        tables_group = RibbonGroup("Tables")
        tables_group.add_large_button(QIcon(), "Table", "Insert table")
        tab.add_group(tables_group)

        # Illustrations group
        illustrations_group = RibbonGroup("Illustrations")
        illustrations_group.add_large_button(QIcon(), "Picture", "Insert picture")
        illustrations_group.add_large_button(QIcon(), "Shapes", "Insert shapes")
        illustrations_group.add_large_button(QIcon(), "Chart", "Insert chart")
        illustrations_group.add_large_button(QIcon(), "SmartArt", "Insert SmartArt")
        tab.add_group(illustrations_group)

        # Links group
        links_group = RibbonGroup("Links")
        links_group.add_large_button(QIcon(), "Link", "Insert hyperlink")
        links_group.add_large_button(QIcon(), "Bookmark", "Insert bookmark")
        tab.add_group(links_group)

        # Header & Footer group
        header_group = RibbonGroup("Header & Footer")
        header_group.add_small_button(QIcon(), "Header", "Edit header")
        header_group.add_small_button(QIcon(), "Footer", "Edit footer")
        header_group.add_small_button(QIcon(), "Page Number", "Insert page number")
        tab.add_group(header_group)

        # Symbols group
        symbols_group = RibbonGroup("Symbols")
        symbols_group.add_large_button(QIcon(), "Equation", "Insert equation")
        symbols_group.add_large_button(QIcon(), "Symbol", "Insert symbol")
        tab.add_group(symbols_group)

        return tab

    def create_design_tab(self) -> RibbonTab:
        """Create the Design tab."""
        tab = RibbonTab("Design")

        # Themes group
        themes_group = RibbonGroup("Themes")
        themes_group.add_large_button(QIcon(), "Themes", "Document themes")
        themes_group.add_small_button(QIcon(), "Colors", "Theme colors")
        themes_group.add_small_button(QIcon(), "Fonts", "Theme fonts")
        tab.add_group(themes_group)

        # Page Background group
        background_group = RibbonGroup("Page Background")
        background_group.add_large_button(QIcon(), "Watermark", "Add watermark")
        background_group.add_large_button(QIcon(), "Page Color", "Page color")
        background_group.add_large_button(QIcon(), "Borders", "Page borders")
        tab.add_group(background_group)

        return tab

    def create_layout_tab(self) -> RibbonTab:
        """Create the Layout tab."""
        tab = RibbonTab("Layout")

        # Page Setup group
        page_setup_group = RibbonGroup("Page Setup")
        page_setup_group.add_small_button(QIcon(), "Margins", "Page margins")
        page_setup_group.add_small_button(QIcon(), "Orientation", "Page orientation")
        page_setup_group.add_small_button(QIcon(), "Size", "Page size")
        page_setup_group.add_small_button(QIcon(), "Columns", "Columns")
        tab.add_group(page_setup_group)

        # Arrange group
        arrange_group = RibbonGroup("Arrange")
        arrange_group.add_small_button(QIcon(), "Position", "Position")
        arrange_group.add_small_button(QIcon(), "Wrap Text", "Text wrapping")
        arrange_group.add_small_button(QIcon(), "Align", "Align objects")
        tab.add_group(arrange_group)

        return tab

    def create_view_tab(self) -> RibbonTab:
        """Create the View tab."""
        tab = RibbonTab("View")

        # Views group
        views_group = RibbonGroup("Views")
        views_group.add_small_button(QIcon(), "Print Layout", "Print layout")
        views_group.add_small_button(QIcon(), "Web Layout", "Web layout")
        views_group.add_small_button(QIcon(), "Draft", "Draft view")
        tab.add_group(views_group)

        # Zoom group
        zoom_group = RibbonGroup("Zoom")
        zoom_group.add_large_button(QIcon(), "Zoom", "Zoom level")
        zoom_group.add_small_button(QIcon(), "100%", "Zoom to 100%")
        zoom_group.add_small_button(QIcon(), "Page Width", "Fit to page width")
        tab.add_group(zoom_group)

        # Window group
        window_group = RibbonGroup("Window")
        window_group.add_small_button(QIcon(), "New Window", "New window")
        window_group.add_small_button(QIcon(), "Split", "Split view")
        window_group.add_small_button(QIcon(), "Side by Side", "View side by side")
        tab.add_group(window_group)

        return tab
