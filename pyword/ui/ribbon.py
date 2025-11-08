"""Ribbon interface for PyWord - modern Microsoft Office-style UI."""

from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QToolButton,
                               QLabel, QFrame, QScrollArea, QSizePolicy, QPushButton,
                               QButtonGroup, QGridLayout, QSpacerItem, QMenu, QStackedWidget,
                               QComboBox, QFontComboBox, QSpinBox, QListWidget, QTextEdit)
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


class BackstageView(QWidget):
    """Microsoft Word-style backstage view shown when File tab is clicked."""

    close_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        """Initialize the backstage view UI."""
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Left sidebar (colored teal/blue, like Word)
        self.sidebar = QWidget()
        self.sidebar.setFixedWidth(220)
        self.sidebar.setStyleSheet("""
            QWidget {
                background-color: #0078D4;
            }
        """)
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)

        # Back button
        back_btn = QPushButton("← Back")
        back_btn.setFlat(True)
        back_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078D4;
                color: white;
                border: none;
                padding: 18px 24px;
                text-align: left;
                font-size: 15px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #106EBE;
            }
            QPushButton:pressed {
                background-color: #005A9E;
            }
        """)
        back_btn.clicked.connect(self.close_requested.emit)
        sidebar_layout.addWidget(back_btn)

        # Menu items
        menu_items = [
            "Info",
            "New",
            "Open",
            "Save",
            "Save As",
            "Print",
            "Share",
            "Export",
            "Close"
        ]

        self.menu_buttons = []
        for item in menu_items:
            btn = QPushButton(item)
            btn.setFlat(True)
            btn.setCheckable(True)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    color: white;
                    border: none;
                    border-left: 3px solid transparent;
                    padding: 14px 24px;
                    text-align: left;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #106EBE;
                    border-left: 3px solid white;
                }
                QPushButton:checked {
                    background-color: #005A9E;
                    border-left: 3px solid white;
                }
            """)
            self.menu_buttons.append(btn)
            sidebar_layout.addWidget(btn)

        sidebar_layout.addStretch()
        main_layout.addWidget(self.sidebar)

        # Right content area
        self.content_area = QWidget()
        self.content_area.setStyleSheet("""
            QWidget {
                background-color: #F3F2F1;
            }
        """)
        content_layout = QVBoxLayout(self.content_area)
        content_layout.setContentsMargins(50, 50, 50, 50)
        content_layout.setSpacing(20)

        # Placeholder content
        title = QLabel("Info")
        title.setStyleSheet("""
            font-size: 28px;
            font-weight: bold;
            color: #323130;
            margin-bottom: 10px;
        """)
        content_layout.addWidget(title)

        # Info cards
        info_sections = [
            ("Protect Document", "Control what types of changes people can make to this document."),
            ("Inspect Document", "Check the document for hidden properties or personal information."),
            ("Manage Document", "Check in, check out, and manage document versions.")
        ]

        for section_title, section_desc in info_sections:
            section_widget = QWidget()
            section_widget.setStyleSheet("""
                QWidget {
                    background-color: white;
                    border-radius: 4px;
                    padding: 20px;
                }
            """)
            section_layout = QVBoxLayout(section_widget)
            section_layout.setContentsMargins(20, 20, 20, 20)

            section_title_label = QLabel(section_title)
            section_title_label.setStyleSheet("font-size: 16px; font-weight: 600; color: #323130;")
            section_layout.addWidget(section_title_label)

            section_desc_label = QLabel(section_desc)
            section_desc_label.setStyleSheet("font-size: 13px; color: #605E5C; margin-top: 8px;")
            section_desc_label.setWordWrap(True)
            section_layout.addWidget(section_desc_label)

            content_layout.addWidget(section_widget)

        content_layout.addStretch()

        main_layout.addWidget(self.content_area, 1)


class RibbonBar(QWidget):
    """Modern ribbon interface bar."""

    tab_changed = Signal(int)
    file_tab_clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.tabs = []
        self.tab_buttons = []
        self.current_tab_index = 0
        self.backstage_view = None
        # Store button references for easy access
        self.buttons = {}
        self.setup_ui()

    def setup_ui(self):
        """Initialize the ribbon UI."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Constrain the ribbon height (Microsoft Word style)
        self.setMaximumHeight(150)

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
            # Check if this is the File tab (index 0)
            if index == 0 and self.tabs[0].title == "File":
                # Emit signal for File tab click
                self.file_tab_clicked.emit()
                # Don't change the visual state - keep previous tab selected
                return

            # Uncheck all buttons
            for btn in self.tab_buttons:
                btn.setChecked(False)

            # Check current button
            self.tab_buttons[index].setChecked(True)

            # Show tab content using stacked widget
            self.stacked_widget.setCurrentIndex(index)
            self.current_tab_index = index
            self.tab_changed.emit(index)

    def show_backstage(self):
        """Show the backstage view."""
        if not self.backstage_view:
            # Create backstage view as a child of the main window
            parent = self.window()
            self.backstage_view = BackstageView(parent)
            self.backstage_view.close_requested.connect(self.hide_backstage)

        # Position it to cover the entire window
        parent = self.window()
        self.backstage_view.setGeometry(0, 0, parent.width(), parent.height())
        self.backstage_view.raise_()
        self.backstage_view.show()

    def hide_backstage(self):
        """Hide the backstage view."""
        if self.backstage_view:
            self.backstage_view.hide()

    def create_file_tab(self) -> RibbonTab:
        """Create the File tab with document operations (backstage view)."""
        tab = RibbonTab("File")

        # Document operations group
        file_group = RibbonGroup("Document")
        self.buttons['new'] = file_group.add_large_button(QIcon(), "New", "Create new document")
        self.buttons['open'] = file_group.add_large_button(QIcon(), "Open", "Open document")
        self.buttons['save'] = file_group.add_large_button(QIcon(), "Save", "Save document")
        self.buttons['save_as'] = file_group.add_large_button(QIcon(), "Save As", "Save document as")
        tab.add_group(file_group)

        # Print group
        print_group = RibbonGroup("Print")
        self.buttons['print'] = print_group.add_large_button(QIcon(), "Print", "Print document")
        self.buttons['print_preview'] = print_group.add_large_button(QIcon(), "Preview", "Print preview")
        tab.add_group(print_group)

        # Share group
        share_group = RibbonGroup("Share")
        self.buttons['export'] = share_group.add_large_button(QIcon(), "Export", "Export document")
        self.buttons['export_pdf'] = share_group.add_small_button(QIcon(), "PDF", "Export as PDF")
        tab.add_group(share_group)

        return tab

    def create_home_tab(self) -> RibbonTab:
        """Create the Home tab with common formatting options."""
        tab = RibbonTab("Home")

        # Clipboard group
        clipboard_group = RibbonGroup("Clipboard")
        # Make Paste button more prominent with larger size
        self.buttons['paste'] = clipboard_group.add_large_button(QIcon(), "Paste", "Paste from clipboard")
        self.buttons['paste'].setIconSize(QSize(36, 36))  # Larger icon
        self.buttons['paste'].setFixedSize(75, 70)  # Slightly larger button
        self.buttons['cut'] = clipboard_group.add_small_button(QIcon(), "Cut", "Cut to clipboard")
        self.buttons['copy'] = clipboard_group.add_small_button(QIcon(), "Copy", "Copy to clipboard")
        self.buttons['format_painter'] = clipboard_group.add_small_button(QIcon(), "Format\nPainter", "Format painter")
        tab.add_group(clipboard_group)

        # Font group (with selectors)
        font_group = RibbonGroup("Font")

        # Row 0: Font selector (changed to Calibri (Body) as default)
        self.buttons['font_family'] = QFontComboBox()
        self.buttons['font_family'].setMaximumWidth(150)
        self.buttons['font_family'].setMinimumWidth(120)
        self.buttons['font_family'].setCurrentFont(QFont("Calibri"))
        self.buttons['font_family'].setStyleSheet("""
            QFontComboBox {
                border: 1px solid #D2D0CE;
                border-radius: 2px;
                padding: 3px;
                background: white;
            }
        """)
        font_group.content_layout.addWidget(self.buttons['font_family'], 0, 0, 1, 2)

        # Row 0: Font size
        self.buttons['font_size'] = QSpinBox()
        self.buttons['font_size'].setRange(8, 72)
        self.buttons['font_size'].setValue(11)
        self.buttons['font_size'].setMaximumWidth(50)
        self.buttons['font_size'].setStyleSheet("""
            QSpinBox {
                border: 1px solid #D2D0CE;
                border-radius: 2px;
                padding: 3px;
                background: white;
            }
        """)
        font_group.content_layout.addWidget(self.buttons['font_size'], 0, 2, 1, 1)

        # Row 0: Font size increase/decrease buttons
        self.buttons['font_size_up'] = QToolButton()
        self.buttons['font_size_up'].setText("A↑")
        self.buttons['font_size_up'].setToolTip("Increase Font Size")
        self.buttons['font_size_up'].setFixedSize(24, 24)
        self.buttons['font_size_up'].setStyleSheet("""
            QToolButton {
                border: 1px solid #D2D0CE;
                border-radius: 2px;
                background: white;
                font-size: 10px;
            }
            QToolButton:hover {
                background-color: #F3F2F1;
            }
        """)
        font_group.content_layout.addWidget(self.buttons['font_size_up'], 0, 3, 1, 1)

        self.buttons['font_size_down'] = QToolButton()
        self.buttons['font_size_down'].setText("A↓")
        self.buttons['font_size_down'].setToolTip("Decrease Font Size")
        self.buttons['font_size_down'].setFixedSize(24, 24)
        self.buttons['font_size_down'].setStyleSheet("""
            QToolButton {
                border: 1px solid #D2D0CE;
                border-radius: 2px;
                background: white;
                font-size: 10px;
            }
            QToolButton:hover {
                background-color: #F3F2F1;
            }
        """)
        font_group.content_layout.addWidget(self.buttons['font_size_down'], 0, 4, 1, 1)

        # Row 1: Bold, Italic, Underline buttons
        font_group.current_col = 0
        font_group.current_row = 1
        self.buttons['bold'] = font_group.add_small_button(QIcon(), "B", "Bold (Ctrl+B)")
        self.buttons['bold'].setFont(QFont("Arial", 10, QFont.Bold))
        self.buttons['bold'].setCheckable(True)

        self.buttons['italic'] = font_group.add_small_button(QIcon(), "I", "Italic (Ctrl+I)")
        self.buttons['italic'].setFont(QFont("Arial", 10, QFont.Normal, True))
        self.buttons['italic'].setCheckable(True)

        # Underline button with dropdown
        self.buttons['underline'] = font_group.add_small_button(QIcon(), "U▾", "Underline (Ctrl+U)")
        underline_font = QFont("Arial", 10)
        underline_font.setUnderline(True)
        self.buttons['underline'].setFont(underline_font)
        self.buttons['underline'].setCheckable(True)

        # Strikethrough button
        self.buttons['strikethrough'] = font_group.add_small_button(QIcon(), "abc", "Strikethrough")
        strikethrough_font = QFont("Arial", 10)
        strikethrough_font.setStrikeOut(True)
        self.buttons['strikethrough'].setFont(strikethrough_font)
        self.buttons['strikethrough'].setCheckable(True)

        # Subscript button
        self.buttons['subscript'] = font_group.add_small_button(QIcon(), "X₂", "Subscript")
        self.buttons['subscript'].setCheckable(True)
        self.buttons['subscript'].setStyleSheet("""
            QToolButton {
                border: 1px solid transparent;
                border-radius: 2px;
                padding: 3px 8px;
                background: transparent;
                color: #323130;
                font-size: 11px;
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

        # Superscript button
        font_group.current_col = 0
        font_group.current_row = 2
        self.buttons['superscript'] = font_group.add_small_button(QIcon(), "X²", "Superscript")
        self.buttons['superscript'].setCheckable(True)
        self.buttons['superscript'].setStyleSheet("""
            QToolButton {
                border: 1px solid transparent;
                border-radius: 2px;
                padding: 3px 8px;
                background: transparent;
                color: #323130;
                font-size: 11px;
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

        # Text Effects button
        self.buttons['text_effects'] = font_group.add_small_button(QIcon(), "A✦", "Text Effects")

        # Text Highlight Color button with dropdown
        self.buttons['highlight_color'] = font_group.add_small_button(QIcon(), "🖍▾", "Text Highlight Color")

        # Font Color button with dropdown
        self.buttons['font_color'] = font_group.add_small_button(QIcon(), "A▾", "Font Color")

        tab.add_group(font_group)

        # Paragraph group
        paragraph_group = RibbonGroup("Paragraph")

        # Row 0: Lists and indents
        self.buttons['bullets'] = paragraph_group.add_small_button(QIcon(), "• ▾", "Bullets")
        self.buttons['numbering'] = paragraph_group.add_small_button(QIcon(), "1. ▾", "Numbering")
        self.buttons['multilevel_list'] = paragraph_group.add_small_button(QIcon(), "≣▾", "Multilevel List")
        self.buttons['decrease_indent'] = paragraph_group.add_small_button(QIcon(), "⬅", "Decrease Indent")
        self.buttons['increase_indent'] = paragraph_group.add_small_button(QIcon(), "➡", "Increase Indent")

        # Row 1: Sorting and show/hide
        paragraph_group.current_col = 0
        paragraph_group.current_row = 1
        self.buttons['sort'] = paragraph_group.add_small_button(QIcon(), "AZ↓", "Sort")
        self.buttons['show_hide'] = paragraph_group.add_small_button(QIcon(), "¶", "Show/Hide ¶")

        # Alignment buttons
        self.buttons['align_left'] = paragraph_group.add_small_button(QIcon(), "≡", "Align Left")
        self.buttons['align_center'] = paragraph_group.add_small_button(QIcon(), "▬", "Center")
        self.buttons['align_right'] = paragraph_group.add_small_button(QIcon(), "≡", "Align Right")

        # Row 2: More paragraph options
        paragraph_group.current_col = 0
        paragraph_group.current_row = 2
        self.buttons['align_justify'] = paragraph_group.add_small_button(QIcon(), "≣", "Justify")
        self.buttons['line_spacing'] = paragraph_group.add_small_button(QIcon(), "↕▾", "Line Spacing")
        self.buttons['shading'] = paragraph_group.add_small_button(QIcon(), "⬜▾", "Shading")
        self.buttons['borders'] = paragraph_group.add_small_button(QIcon(), "▦▾", "Borders")

        tab.add_group(paragraph_group)

        # Styles group
        styles_group = RibbonGroup("Styles")

        # Create style preview boxes
        style_layout = QHBoxLayout()
        style_layout.setSpacing(2)

        # Add some common styles
        styles = [
            ("Normal", "Arial", 11, False, False),
            ("Heading 1", "Arial", 16, True, False),
            ("Heading 2", "Arial", 13, True, False),
            ("Title", "Arial", 26, True, False),
        ]

        for style_name, font_name, font_size, bold, italic in styles:
            style_btn = QPushButton(style_name)
            style_btn.setMinimumWidth(70)
            style_btn.setMaximumHeight(60)
            font = QFont(font_name, min(font_size, 11))  # Scale down for preview
            font.setBold(bold)
            font.setItalic(italic)
            style_btn.setFont(font)
            style_btn.setStyleSheet("""
                QPushButton {
                    border: 1px solid #D2D0CE;
                    border-radius: 2px;
                    padding: 8px 4px;
                    background: white;
                    text-align: center;
                }
                QPushButton:hover {
                    background-color: #F3F2F1;
                    border: 1px solid #0078D4;
                }
            """)
            style_layout.addWidget(style_btn)

        # Add the style layout to the group
        style_container = QWidget()
        style_container.setLayout(style_layout)
        styles_group.content_layout.addWidget(style_container, 0, 0, 3, 4)

        tab.add_group(styles_group)

        # Editing group
        editing_group = RibbonGroup("Editing")
        self.buttons['find'] = editing_group.add_large_button(QIcon(), "Find", "Find text")
        self.buttons['replace'] = editing_group.add_small_button(QIcon(), "Replace", "Find and replace")
        self.buttons['select_all'] = editing_group.add_small_button(QIcon(), "Select", "Select text")
        tab.add_group(editing_group)

        return tab

    def create_insert_tab(self) -> RibbonTab:
        """Create the Insert tab."""
        tab = RibbonTab("Insert")

        # Tables group
        tables_group = RibbonGroup("Tables")
        self.buttons['insert_table'] = tables_group.add_large_button(QIcon(), "Table", "Insert table")
        tab.add_group(tables_group)

        # Illustrations group
        illustrations_group = RibbonGroup("Illustrations")
        self.buttons['insert_picture'] = illustrations_group.add_large_button(QIcon(), "Picture", "Insert picture")
        self.buttons['insert_shapes'] = illustrations_group.add_large_button(QIcon(), "Shapes", "Insert shapes")
        self.buttons['insert_chart'] = illustrations_group.add_large_button(QIcon(), "Chart", "Insert chart")
        self.buttons['insert_smartart'] = illustrations_group.add_large_button(QIcon(), "SmartArt", "Insert SmartArt")
        tab.add_group(illustrations_group)

        # Links group
        links_group = RibbonGroup("Links")
        self.buttons['insert_link'] = links_group.add_large_button(QIcon(), "Link", "Insert hyperlink")
        self.buttons['insert_bookmark'] = links_group.add_large_button(QIcon(), "Bookmark", "Insert bookmark")
        tab.add_group(links_group)

        # Header & Footer group
        header_group = RibbonGroup("Header & Footer")
        self.buttons['insert_header'] = header_group.add_small_button(QIcon(), "Header", "Edit header")
        self.buttons['insert_footer'] = header_group.add_small_button(QIcon(), "Footer", "Edit footer")
        self.buttons['insert_page_number'] = header_group.add_small_button(QIcon(), "Page Number", "Insert page number")
        tab.add_group(header_group)

        # Symbols group
        symbols_group = RibbonGroup("Symbols")
        self.buttons['insert_equation'] = symbols_group.add_large_button(QIcon(), "Equation", "Insert equation")
        self.buttons['insert_symbol'] = symbols_group.add_large_button(QIcon(), "Symbol", "Insert symbol")
        tab.add_group(symbols_group)

        return tab

    def create_design_tab(self) -> RibbonTab:
        """Create the Design tab."""
        tab = RibbonTab("Design")

        # Themes group
        themes_group = RibbonGroup("Themes")
        self.buttons['design_themes'] = themes_group.add_large_button(QIcon(), "Themes", "Document themes")
        self.buttons['design_colors'] = themes_group.add_small_button(QIcon(), "Colors", "Theme colors")
        self.buttons['design_fonts'] = themes_group.add_small_button(QIcon(), "Fonts", "Theme fonts")
        tab.add_group(themes_group)

        # Page Background group
        background_group = RibbonGroup("Page Background")
        self.buttons['design_watermark'] = background_group.add_large_button(QIcon(), "Watermark", "Add watermark")
        self.buttons['design_page_color'] = background_group.add_large_button(QIcon(), "Page Color", "Page color")
        self.buttons['design_page_borders'] = background_group.add_large_button(QIcon(), "Borders", "Page borders")
        tab.add_group(background_group)

        return tab

    def create_layout_tab(self) -> RibbonTab:
        """Create the Layout tab."""
        tab = RibbonTab("Layout")

        # Page Setup group
        page_setup_group = RibbonGroup("Page Setup")
        self.buttons['layout_margins'] = page_setup_group.add_small_button(QIcon(), "Margins", "Page margins")
        self.buttons['layout_orientation'] = page_setup_group.add_small_button(QIcon(), "Orientation", "Page orientation")
        self.buttons['layout_size'] = page_setup_group.add_small_button(QIcon(), "Size", "Page size")
        self.buttons['layout_columns'] = page_setup_group.add_small_button(QIcon(), "Columns", "Columns")
        tab.add_group(page_setup_group)

        # Arrange group
        arrange_group = RibbonGroup("Arrange")
        self.buttons['layout_position'] = arrange_group.add_small_button(QIcon(), "Position", "Position")
        self.buttons['layout_wrap_text'] = arrange_group.add_small_button(QIcon(), "Wrap Text", "Text wrapping")
        self.buttons['layout_align'] = arrange_group.add_small_button(QIcon(), "Align", "Align objects")
        tab.add_group(arrange_group)

        return tab

    def create_view_tab(self) -> RibbonTab:
        """Create the View tab."""
        tab = RibbonTab("View")

        # Views group
        views_group = RibbonGroup("Views")
        self.buttons['view_print_layout'] = views_group.add_small_button(QIcon(), "Print Layout", "Print layout")
        self.buttons['view_web_layout'] = views_group.add_small_button(QIcon(), "Web Layout", "Web layout")
        self.buttons['view_draft'] = views_group.add_small_button(QIcon(), "Draft", "Draft view")
        tab.add_group(views_group)

        # Zoom group
        zoom_group = RibbonGroup("Zoom")
        self.buttons['view_zoom'] = zoom_group.add_large_button(QIcon(), "Zoom", "Zoom level")
        self.buttons['view_zoom_100'] = zoom_group.add_small_button(QIcon(), "100%", "Zoom to 100%")
        self.buttons['view_page_width'] = zoom_group.add_small_button(QIcon(), "Page Width", "Fit to page width")
        tab.add_group(zoom_group)

        # Window group
        window_group = RibbonGroup("Window")
        self.buttons['view_new_window'] = window_group.add_small_button(QIcon(), "New Window", "New window")
        self.buttons['view_split'] = window_group.add_small_button(QIcon(), "Split", "Split view")
        self.buttons['view_side_by_side'] = window_group.add_small_button(QIcon(), "Side by Side", "View side by side")
        tab.add_group(window_group)

        return tab
