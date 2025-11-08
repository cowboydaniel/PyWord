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
        """Initialize the tab UI (compact)."""
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(4, 4, 4, 4)
        self.layout.setSpacing(8)  # Tighter spacing between groups
        self.layout.addStretch()

    def add_group(self, group: 'RibbonGroup'):
        """Add a group to this tab."""
        # Add separator before group (except for first group) - subtle Word-style divider
        if len(self.groups) > 0:
            separator = QFrame()
            separator.setFrameShape(QFrame.VLine)
            separator.setFrameShadow(QFrame.Plain)
            separator.setFixedWidth(1)
            separator.setStyleSheet("""
                QFrame {
                    background-color: #EDEBE9;
                    border: none;
                    margin: 4px 0px;
                }
            """)
            # Insert separator before the stretch at the end
            self.layout.insertWidget(self.layout.count() - 1, separator)

        # Add the group before the stretch at the end
        self.groups.append(group)
        self.layout.insertWidget(self.layout.count() - 1, group)


class RibbonGroup(QWidget):
    """A group of related controls in a ribbon tab."""

    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.title = title
        self.buttons = []
        self.setup_ui()

    def setup_ui(self):
        """Initialize the group UI (compact Microsoft Word style)."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(6, 2, 6, 2)
        main_layout.setSpacing(1)

        # Content area
        self.content_layout = QGridLayout()
        self.content_layout.setSpacing(3)  # Tighter spacing for compact design
        main_layout.addLayout(self.content_layout)

        # Group title (Microsoft Word style - small, centered, subtle)
        title_label = QLabel(self.title)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("""
            font-size: 10px;
            color: #605E5C;
            font-weight: normal;
            padding-top: 1px;
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

        # Constrain the ribbon height (Microsoft Word style - compact)
        self.setMaximumHeight(140)

        # Tab bar (Word-like theme with subtle gradient)
        self.tab_bar_widget = QWidget()
        self.tab_bar_widget.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FAFAFA, stop:1 #F3F2F1);
                border-bottom: 1px solid #C6C6C6;
            }
        """)
        self.tab_bar_layout = QHBoxLayout(self.tab_bar_widget)
        self.tab_bar_layout.setContentsMargins(8, 0, 8, 0)
        self.tab_bar_layout.setSpacing(2)

        # Add spacer before right-side controls
        self.tab_bar_layout.addStretch()

        # Add search box (Tell me what you want to do...) - Word-like theme
        from PySide6.QtWidgets import QLineEdit
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Tell me what you want to do...")
        self.search_box.setFixedWidth(250)
        self.search_box.setFixedHeight(24)
        self.search_box.setStyleSheet("""
            QLineEdit {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FFFFFF, stop:1 #F7F7F7);
                color: #323130;
                border: 1px solid #ABABAB;
                border-radius: 2px;
                padding: 4px 8px;
                font-size: 12px;
            }
            QLineEdit:hover {
                border: 1px solid #7A7A7A;
                background: white;
            }
            QLineEdit:focus {
                border: 1px solid #0078D4;
                background: white;
            }
        """)
        self.tab_bar_layout.addWidget(self.search_box)

        # Add Share button
        self.share_button = QPushButton("Share")
        self.share_button.setFixedHeight(24)
        self.share_button.setStyleSheet("""
            QPushButton {
                background-color: #0078D4;
                color: #FFFFFF;
                border: 1px solid #0078D4;
                border-radius: 2px;
                padding: 4px 12px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #106EBE;
                border: 1px solid #106EBE;
            }
            QPushButton:pressed {
                background-color: #005A9E;
                border: 1px solid #005A9E;
            }
        """)
        self.tab_bar_layout.addWidget(self.share_button)

        # Add User icon/button
        self.user_button = QPushButton("👤")
        self.user_button.setFixedSize(32, 24)
        self.user_button.setToolTip("Account")
        self.user_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #605E5C;
                border: 1px solid transparent;
                border-radius: 2px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #E1DFDD;
                border: 1px solid #D2D0CE;
            }
        """)
        self.tab_bar_layout.addWidget(self.user_button)

        main_layout.addWidget(self.tab_bar_widget)

        # Content area with stacked widget for tabs (Word-like theme)
        self.content_widget = QWidget()
        self.content_widget.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FFFFFF, stop:1 #FAFAFA);
                border-bottom: 1px solid #C6C6C6;
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
        self.scroll_area.setMaximumHeight(110)  # Limit ribbon height (compact)
        self.content_layout.addWidget(self.scroll_area)

    def add_tab(self, tab: RibbonTab) -> int:
        """Add a tab to the ribbon."""
        self.tabs.append(tab)

        # Add tab to stacked widget
        self.stacked_widget.addWidget(tab)

        # Create tab button (Word-like theme with subtle gradients)
        button = QPushButton(tab.title)
        button.setCheckable(True)
        button.setFlat(True)
        button.setStyleSheet("""
            QPushButton {
                padding: 6px 18px;
                border: none;
                border-bottom: 3px solid transparent;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FAFAFA, stop:1 #F3F2F1);
                font-weight: normal;
                font-size: 13px;
                color: #323130;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #F0F0F0, stop:1 #E5E5E5);
            }
            QPushButton:checked {
                border-bottom: 3px solid #0078D4;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FFFFFF, stop:1 #FAFAFA);
                color: #0078D4;
                font-weight: 500;
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
        """Create the File tab (shows backstage view, no ribbon content)."""
        tab = RibbonTab("File")

        # Note: File tab has no content - clicking it shows the backstage view
        # This matches Microsoft Word behavior where File tab only triggers backstage

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

        # Compact spacing (Word-like)
        font_group.content_layout.setHorizontalSpacing(4)
        font_group.content_layout.setVerticalSpacing(2)

        # Row 0: Font selector (changed to Calibri (Body) as default - Word style)
        self.buttons['font_family'] = QFontComboBox()
        self.buttons['font_family'].setMaximumWidth(200)  # Wider to match Word
        self.buttons['font_family'].setMinimumWidth(180)
        self.buttons['font_family'].setCurrentFont(QFont("Calibri"))
        # Set custom display text to show "Calibri (Body)" like Word
        self.buttons['font_family'].setEditable(True)
        self.buttons['font_family'].lineEdit().setText("Calibri (Body)")
        self.buttons['font_family'].setEditable(False)
        self.buttons['font_family'].setStyleSheet("""
            QFontComboBox {
                border: 1px solid #ABABAB;
                border-radius: 2px;
                padding: 3px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FFFFFF, stop:1 #F7F7F7);
                color: #323130;
            }
            QFontComboBox:hover {
                border: 1px solid #C6C6C6;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FFFFFF, stop:1 #FAFAFA);
            }
            QFontComboBox:focus {
                border: 1px solid #0078D4;
                background: white;
            }
        """)
        font_group.content_layout.addWidget(self.buttons['font_family'], 0, 0, 1, 2)

        # Row 0: Font size (Word-like theme)
        self.buttons['font_size'] = QSpinBox()
        self.buttons['font_size'].setRange(8, 72)
        self.buttons['font_size'].setValue(11)
        self.buttons['font_size'].setMaximumWidth(50)
        self.buttons['font_size'].setStyleSheet("""
            QSpinBox {
                border: 1px solid #ABABAB;
                border-radius: 2px;
                padding: 3px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FFFFFF, stop:1 #F7F7F7);
                color: #323130;
            }
            QSpinBox:hover {
                border: 1px solid #C6C6C6;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FFFFFF, stop:1 #FAFAFA);
            }
            QSpinBox:focus {
                border: 1px solid #0078D4;
                background: white;
            }
        """)
        font_group.content_layout.addWidget(self.buttons['font_size'], 0, 2, 1, 1)

        # Row 0: Font size increase/decrease buttons (Word-like theme)
        self.buttons['font_size_up'] = QToolButton()
        self.buttons['font_size_up'].setText("A↑")
        self.buttons['font_size_up'].setToolTip("Increase Font Size")
        self.buttons['font_size_up'].setFixedSize(24, 24)
        self.buttons['font_size_up'].setStyleSheet("""
            QToolButton {
                border: 1px solid #ABABAB;
                border-radius: 2px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FFFFFF, stop:1 #F7F7F7);
                font-size: 10px;
                color: #323130;
            }
            QToolButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FAFAFA, stop:1 #ECECEC);
                border: 1px solid #C6C6C6;
            }
            QToolButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #E5E5E5, stop:1 #F0F0F0);
                border: 1px solid #ADADAD;
            }
        """)
        font_group.content_layout.addWidget(self.buttons['font_size_up'], 0, 3, 1, 1)

        self.buttons['font_size_down'] = QToolButton()
        self.buttons['font_size_down'].setText("A↓")
        self.buttons['font_size_down'].setToolTip("Decrease Font Size")
        self.buttons['font_size_down'].setFixedSize(24, 24)
        self.buttons['font_size_down'].setStyleSheet("""
            QToolButton {
                border: 1px solid #ABABAB;
                border-radius: 2px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FFFFFF, stop:1 #F7F7F7);
                font-size: 10px;
                color: #323130;
            }
            QToolButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FAFAFA, stop:1 #ECECEC);
                border: 1px solid #C6C6C6;
            }
            QToolButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #E5E5E5, stop:1 #F0F0F0);
                border: 1px solid #ADADAD;
            }
        """)
        font_group.content_layout.addWidget(self.buttons['font_size_down'], 0, 4, 1, 1)

        # Create common style for formatting buttons (Word-like theme)
        button_style = """
            QToolButton {
                border: 1px solid transparent;
                border-radius: 2px;
                padding: 2px 6px;
                background: transparent;
                color: #323130;
                font-size: 12px;
                text-align: center;
                min-width: 28px;
            }
            QToolButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FAFAFA, stop:1 #F0F0F0);
                border: 1px solid #C6C6C6;
                box-shadow: 0px 1px 2px rgba(0, 0, 0, 0.05);
            }
            QToolButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #E5E5E5, stop:1 #F0F0F0);
                border: 1px solid #ADADAD;
            }
            QToolButton:checked {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #E5F3FB, stop:1 #D0E9F7);
                border: 1px solid #92C0E0;
            }
        """

        # Row 1: Bold button
        self.buttons['bold'] = QToolButton()
        self.buttons['bold'].setText("B")
        self.buttons['bold'].setToolTip("Bold (Ctrl+B)")
        self.buttons['bold'].setFont(QFont("Arial", 10, QFont.Bold))
        self.buttons['bold'].setCheckable(True)
        self.buttons['bold'].setFixedHeight(24)
        self.buttons['bold'].setStyleSheet(button_style)
        font_group.content_layout.addWidget(self.buttons['bold'], 1, 0)

        # Row 1: Italic button
        self.buttons['italic'] = QToolButton()
        self.buttons['italic'].setText("I")
        self.buttons['italic'].setToolTip("Italic (Ctrl+I)")
        italic_font = QFont("Arial", 10)
        italic_font.setItalic(True)
        self.buttons['italic'].setFont(italic_font)
        self.buttons['italic'].setCheckable(True)
        self.buttons['italic'].setFixedHeight(24)
        self.buttons['italic'].setStyleSheet(button_style)
        font_group.content_layout.addWidget(self.buttons['italic'], 1, 1)

        # Row 1: Underline button with dropdown
        self.buttons['underline'] = QToolButton()
        self.buttons['underline'].setText("U▾")
        self.buttons['underline'].setToolTip("Underline (Ctrl+U)")
        underline_font = QFont("Arial", 10)
        underline_font.setUnderline(True)
        self.buttons['underline'].setFont(underline_font)
        self.buttons['underline'].setCheckable(True)
        self.buttons['underline'].setFixedHeight(24)
        self.buttons['underline'].setStyleSheet(button_style)
        font_group.content_layout.addWidget(self.buttons['underline'], 1, 2)

        # Row 1: Strikethrough button
        self.buttons['strikethrough'] = QToolButton()
        self.buttons['strikethrough'].setText("abc")
        self.buttons['strikethrough'].setToolTip("Strikethrough")
        strikethrough_font = QFont("Arial", 10)
        strikethrough_font.setStrikeOut(True)
        self.buttons['strikethrough'].setFont(strikethrough_font)
        self.buttons['strikethrough'].setCheckable(True)
        self.buttons['strikethrough'].setFixedHeight(24)
        self.buttons['strikethrough'].setStyleSheet(button_style)
        font_group.content_layout.addWidget(self.buttons['strikethrough'], 1, 3)

        # Row 1: Subscript button
        self.buttons['subscript'] = QToolButton()
        self.buttons['subscript'].setText("X₂")
        self.buttons['subscript'].setToolTip("Subscript")
        self.buttons['subscript'].setCheckable(True)
        self.buttons['subscript'].setFixedHeight(24)
        self.buttons['subscript'].setStyleSheet(button_style)
        font_group.content_layout.addWidget(self.buttons['subscript'], 1, 4)

        # Row 2: Superscript button
        self.buttons['superscript'] = QToolButton()
        self.buttons['superscript'].setText("X²")
        self.buttons['superscript'].setToolTip("Superscript")
        self.buttons['superscript'].setCheckable(True)
        self.buttons['superscript'].setFixedHeight(24)
        self.buttons['superscript'].setStyleSheet(button_style)
        font_group.content_layout.addWidget(self.buttons['superscript'], 2, 0)

        # Row 2: Text Effects button
        self.buttons['text_effects'] = QToolButton()
        self.buttons['text_effects'].setText("A✦")
        self.buttons['text_effects'].setToolTip("Text Effects")
        self.buttons['text_effects'].setFixedHeight(24)
        self.buttons['text_effects'].setStyleSheet(button_style)
        font_group.content_layout.addWidget(self.buttons['text_effects'], 2, 1)

        # Row 2: Text Highlight Color button with dropdown (yellow highlight) - Word-like theme
        self.buttons['highlight_color'] = QToolButton()
        self.buttons['highlight_color'].setText("ab")
        self.buttons['highlight_color'].setToolTip("Text Highlight Color")
        self.buttons['highlight_color'].setFixedHeight(24)
        self.buttons['highlight_color'].setStyleSheet("""
            QToolButton {
                border: 1px solid transparent;
                border-radius: 2px;
                padding: 2px 6px;
                background: transparent;
                color: #323130;
                font-size: 12px;
                text-align: center;
                font-weight: bold;
                min-width: 28px;
                background-image: linear-gradient(transparent 65%, #FFED4E 65%);
            }
            QToolButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FAFAFA, stop:0.65 #F0F0F0,
                    stop:0.65 #FFED4E, stop:1 #FFED4E);
                border: 1px solid #C6C6C6;
            }
            QToolButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #E5E5E5, stop:0.65 #F0F0F0,
                    stop:0.65 #FFE838, stop:1 #FFE838);
                border: 1px solid #ADADAD;
            }
        """)
        font_group.content_layout.addWidget(self.buttons['highlight_color'], 2, 2)

        # Row 2: Font Color button with dropdown (red underline) - Word-like theme
        self.buttons['font_color'] = QToolButton()
        self.buttons['font_color'].setText("A")
        self.buttons['font_color'].setToolTip("Font Color")
        self.buttons['font_color'].setFixedHeight(24)
        self.buttons['font_color'].setStyleSheet("""
            QToolButton {
                border: 1px solid transparent;
                border-radius: 2px;
                padding: 2px 6px;
                background: transparent;
                color: #323130;
                font-size: 14px;
                text-align: center;
                font-weight: bold;
                min-width: 28px;
                border-bottom: 3px solid #D13438;
            }
            QToolButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FAFAFA, stop:1 #F0F0F0);
                border: 1px solid #C6C6C6;
                border-bottom: 3px solid #D13438;
            }
            QToolButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #E5E5E5, stop:1 #F0F0F0);
                border: 1px solid #ADADAD;
                border-bottom: 3px solid #D13438;
            }
        """)
        font_group.content_layout.addWidget(self.buttons['font_color'], 2, 3)

        tab.add_group(font_group)

        # Paragraph group
        paragraph_group = RibbonGroup("Paragraph")

        # Compact spacing (Word-like)
        paragraph_group.content_layout.setHorizontalSpacing(4)
        paragraph_group.content_layout.setVerticalSpacing(2)

        # Create common style for paragraph buttons (Word-like theme)
        para_button_style = """
            QToolButton {
                border: 1px solid transparent;
                border-radius: 2px;
                padding: 2px 6px;
                background: transparent;
                color: #323130;
                font-size: 12px;
                text-align: center;
                min-width: 28px;
            }
            QToolButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FAFAFA, stop:1 #F0F0F0);
                border: 1px solid #C6C6C6;
                box-shadow: 0px 1px 2px rgba(0, 0, 0, 0.05);
            }
            QToolButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #E5E5E5, stop:1 #F0F0F0);
                border: 1px solid #ADADAD;
            }
            QToolButton:checked {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #E5F3FB, stop:1 #D0E9F7);
                border: 1px solid #92C0E0;
            }
        """

        # Row 0: Lists and indents
        self.buttons['bullets'] = QToolButton()
        self.buttons['bullets'].setText("• ▾")
        self.buttons['bullets'].setToolTip("Bullets")
        self.buttons['bullets'].setFixedHeight(24)
        self.buttons['bullets'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['bullets'], 0, 0)

        self.buttons['numbering'] = QToolButton()
        self.buttons['numbering'].setText("1. ▾")
        self.buttons['numbering'].setToolTip("Numbering")
        self.buttons['numbering'].setFixedHeight(24)
        self.buttons['numbering'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['numbering'], 0, 1)

        self.buttons['multilevel_list'] = QToolButton()
        self.buttons['multilevel_list'].setText("≣▾")
        self.buttons['multilevel_list'].setToolTip("Multilevel List")
        self.buttons['multilevel_list'].setFixedHeight(24)
        self.buttons['multilevel_list'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['multilevel_list'], 0, 2)

        self.buttons['decrease_indent'] = QToolButton()
        self.buttons['decrease_indent'].setText("⬅")
        self.buttons['decrease_indent'].setToolTip("Decrease Indent")
        self.buttons['decrease_indent'].setFixedHeight(24)
        self.buttons['decrease_indent'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['decrease_indent'], 0, 3)

        self.buttons['increase_indent'] = QToolButton()
        self.buttons['increase_indent'].setText("➡")
        self.buttons['increase_indent'].setToolTip("Increase Indent")
        self.buttons['increase_indent'].setFixedHeight(24)
        self.buttons['increase_indent'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['increase_indent'], 0, 4)

        # Row 1: Sorting, show/hide, and alignment buttons
        self.buttons['sort'] = QToolButton()
        self.buttons['sort'].setText("AZ↓")
        self.buttons['sort'].setToolTip("Sort")
        self.buttons['sort'].setFixedHeight(24)
        self.buttons['sort'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['sort'], 1, 0)

        self.buttons['show_hide'] = QToolButton()
        self.buttons['show_hide'].setText("¶")
        self.buttons['show_hide'].setToolTip("Show/Hide ¶")
        self.buttons['show_hide'].setFixedHeight(24)
        self.buttons['show_hide'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['show_hide'], 1, 1)

        self.buttons['align_left'] = QToolButton()
        self.buttons['align_left'].setText("≡")
        self.buttons['align_left'].setToolTip("Align Left")
        self.buttons['align_left'].setFixedHeight(24)
        self.buttons['align_left'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['align_left'], 1, 2)

        self.buttons['align_center'] = QToolButton()
        self.buttons['align_center'].setText("▬")
        self.buttons['align_center'].setToolTip("Center")
        self.buttons['align_center'].setFixedHeight(24)
        self.buttons['align_center'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['align_center'], 1, 3)

        self.buttons['align_right'] = QToolButton()
        self.buttons['align_right'].setText("≡")
        self.buttons['align_right'].setToolTip("Align Right")
        self.buttons['align_right'].setFixedHeight(24)
        self.buttons['align_right'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['align_right'], 1, 4)

        # Row 2: More paragraph options
        self.buttons['align_justify'] = QToolButton()
        self.buttons['align_justify'].setText("≣")
        self.buttons['align_justify'].setToolTip("Justify")
        self.buttons['align_justify'].setFixedHeight(24)
        self.buttons['align_justify'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['align_justify'], 2, 0)

        self.buttons['line_spacing'] = QToolButton()
        self.buttons['line_spacing'].setText("↕▾")
        self.buttons['line_spacing'].setToolTip("Line Spacing")
        self.buttons['line_spacing'].setFixedHeight(24)
        self.buttons['line_spacing'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['line_spacing'], 2, 1)

        self.buttons['shading'] = QToolButton()
        self.buttons['shading'].setText("⬜▾")
        self.buttons['shading'].setToolTip("Shading")
        self.buttons['shading'].setFixedHeight(24)
        self.buttons['shading'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['shading'], 2, 2)

        self.buttons['borders'] = QToolButton()
        self.buttons['borders'].setText("▦▾")
        self.buttons['borders'].setToolTip("Borders")
        self.buttons['borders'].setFixedHeight(24)
        self.buttons['borders'].setStyleSheet(para_button_style)
        paragraph_group.content_layout.addWidget(self.buttons['borders'], 2, 3)

        tab.add_group(paragraph_group)

        # Styles group (Microsoft Word style - horizontal with dropdown)
        styles_group = RibbonGroup("Styles")

        # Create style preview boxes - showing only 3 main styles like Word
        style_layout = QHBoxLayout()
        style_layout.setSpacing(2)

        # Add Microsoft Word default styles (only 3 visible by default)
        # Style format: (name, font_name, font_size, bold, color)
        styles = [
            ("Normal", "Calibri", 11, False, "#000000"),
            ("No Spacing", "Calibri", 11, False, "#000000"),
            ("Heading 1 ▾", "Calibri Light", 16, False, "#2E75B5"),  # Light blue, larger
        ]

        for style_name, font_name, font_size, bold, color in styles:
            style_btn = QPushButton(style_name)
            style_btn.setMinimumWidth(90)  # Slightly wider for dropdown arrow
            style_btn.setMaximumHeight(60)

            # Create font with specific styling
            font = QFont(font_name if font_name == "Calibri Light" else "Calibri", 11)  # Consistent button size
            font.setBold(bold)
            if font_name == "Calibri Light":
                font.setWeight(QFont.Light)

            style_btn.setFont(font)

            # Add color styling for text
            text_color = color
            style_btn.setStyleSheet(f"""
                QPushButton {{
                    border: 1px solid #D2D0CE;
                    border-radius: 2px;
                    padding: 8px 6px;
                    background: white;
                    text-align: left;
                    color: {text_color};
                }}
                QPushButton:hover {{
                    background-color: #F3F2F1;
                    border: 1px solid #0078D4;
                }}
            """)
            style_layout.addWidget(style_btn)

        # Add "More" dropdown button
        more_styles_btn = QPushButton("⌄")
        more_styles_btn.setToolTip("More Styles")
        more_styles_btn.setMaximumWidth(20)
        more_styles_btn.setMaximumHeight(60)
        more_styles_btn.setStyleSheet("""
            QPushButton {
                border: 1px solid #D2D0CE;
                border-radius: 2px;
                padding: 4px;
                background: white;
            }
            QPushButton:hover {
                background-color: #F3F2F1;
                border: 1px solid #0078D4;
            }
        """)
        style_layout.addWidget(more_styles_btn)

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

    def create_references_tab(self) -> RibbonTab:
        """Create the References tab."""
        tab = RibbonTab("References")

        # Table of Contents group
        toc_group = RibbonGroup("Table of Contents")
        self.buttons['insert_toc'] = toc_group.add_large_button(QIcon(), "Table of\nContents", "Insert table of contents")
        self.buttons['update_toc'] = toc_group.add_small_button(QIcon(), "Update Table", "Update table of contents")
        self.buttons['add_text'] = toc_group.add_small_button(QIcon(), "Add Text", "Add text to table of contents")
        tab.add_group(toc_group)

        # Footnotes group
        footnotes_group = RibbonGroup("Footnotes")
        self.buttons['insert_footnote'] = footnotes_group.add_large_button(QIcon(), "Insert\nFootnote", "Insert footnote")
        self.buttons['insert_endnote'] = footnotes_group.add_large_button(QIcon(), "Insert\nEndnote", "Insert endnote")
        tab.add_group(footnotes_group)

        # Citations & Bibliography group
        citations_group = RibbonGroup("Citations & Bibliography")
        self.buttons['insert_citation'] = citations_group.add_small_button(QIcon(), "Insert Citation", "Insert citation")
        self.buttons['manage_sources'] = citations_group.add_small_button(QIcon(), "Manage Sources", "Manage sources")
        self.buttons['bibliography'] = citations_group.add_small_button(QIcon(), "Bibliography", "Insert bibliography")
        tab.add_group(citations_group)

        # Captions group
        captions_group = RibbonGroup("Captions")
        self.buttons['insert_caption'] = captions_group.add_small_button(QIcon(), "Insert Caption", "Insert caption")
        self.buttons['insert_table_figures'] = captions_group.add_small_button(QIcon(), "Table of Figures", "Insert table of figures")
        self.buttons['cross_reference'] = captions_group.add_small_button(QIcon(), "Cross-reference", "Insert cross-reference")
        tab.add_group(captions_group)

        return tab

    def create_mailings_tab(self) -> RibbonTab:
        """Create the Mailings tab."""
        tab = RibbonTab("Mailings")

        # Create group
        create_group = RibbonGroup("Create")
        self.buttons['envelopes'] = create_group.add_large_button(QIcon(), "Envelopes", "Create envelopes")
        self.buttons['labels'] = create_group.add_large_button(QIcon(), "Labels", "Create labels")
        tab.add_group(create_group)

        # Start Mail Merge group
        mail_merge_group = RibbonGroup("Start Mail Merge")
        self.buttons['start_mail_merge'] = mail_merge_group.add_large_button(QIcon(), "Start Mail\nMerge", "Start mail merge")
        self.buttons['select_recipients'] = mail_merge_group.add_small_button(QIcon(), "Select Recipients", "Select recipients")
        self.buttons['edit_recipient_list'] = mail_merge_group.add_small_button(QIcon(), "Edit Recipients", "Edit recipient list")
        tab.add_group(mail_merge_group)

        # Write & Insert Fields group
        fields_group = RibbonGroup("Write & Insert Fields")
        self.buttons['insert_merge_field'] = fields_group.add_small_button(QIcon(), "Insert Merge Field", "Insert merge field")
        self.buttons['rules'] = fields_group.add_small_button(QIcon(), "Rules", "Mail merge rules")
        self.buttons['match_fields'] = fields_group.add_small_button(QIcon(), "Match Fields", "Match fields")
        tab.add_group(fields_group)

        # Preview Results group
        preview_group = RibbonGroup("Preview Results")
        self.buttons['preview_results'] = preview_group.add_large_button(QIcon(), "Preview\nResults", "Preview merge results")
        tab.add_group(preview_group)

        # Finish group
        finish_group = RibbonGroup("Finish")
        self.buttons['finish_merge'] = finish_group.add_large_button(QIcon(), "Finish &\nMerge", "Finish and merge")
        tab.add_group(finish_group)

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
