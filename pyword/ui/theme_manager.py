"""Theme Manager for PyWord - supports dark and light themes."""

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QObject, Signal, QSettings
from PySide6.QtGui import QPalette, QColor
from enum import Enum


class Theme(Enum):
    """Available themes."""
    LIGHT = "light"
    DARK = "dark"
    AUTO = "auto"  # Follow system theme


class ThemeManager(QObject):
    """Manages application themes (dark/light mode)."""

    theme_changed = Signal(Theme)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings = QSettings("PyWord", "Theme")
        self.current_theme = Theme.LIGHT
        self.load_theme_preference()

    def load_theme_preference(self):
        """Load saved theme preference."""
        saved_theme = self.settings.value("theme", Theme.LIGHT.value)
        try:
            self.current_theme = Theme(saved_theme)
        except ValueError:
            self.current_theme = Theme.LIGHT

    def save_theme_preference(self):
        """Save current theme preference."""
        self.settings.setValue("theme", self.current_theme.value)

    def set_theme(self, theme: Theme):
        """Set the application theme."""
        self.current_theme = theme
        self.save_theme_preference()
        self.apply_theme()
        self.theme_changed.emit(theme)

    def apply_theme(self):
        """Apply the current theme to the application."""
        app = QApplication.instance()
        if not app:
            return

        if self.current_theme == Theme.DARK:
            self.apply_dark_theme(app)
        elif self.current_theme == Theme.LIGHT:
            self.apply_light_theme(app)
        elif self.current_theme == Theme.AUTO:
            # In a full implementation, this would check system theme
            self.apply_light_theme(app)

    def apply_light_theme(self, app: QApplication):
        """Apply light theme to the application (Microsoft Word style)."""
        # Reset to default palette
        app.setStyleSheet("")
        app.setPalette(app.style().standardPalette())

        # Microsoft Word 2019/365 inspired theme stylesheet
        stylesheet = """
            QMainWindow {
                background-color: #F3F2F1;
            }

            QWidget#centralWidget {
                background-color: #F3F2F1;
            }

            QMenuBar {
                background-color: #F3F2F1;
                color: #323130;
                border: none;
                padding: 2px;
            }

            QMenuBar::item {
                background-color: transparent;
                padding: 6px 12px;
                border-radius: 2px;
            }

            QMenuBar::item:selected {
                background-color: #E1DFDD;
            }

            QMenuBar::item:pressed {
                background-color: #D2D0CE;
            }

            QMenu {
                background-color: #FFFFFF;
                color: #323130;
                border: 1px solid #C8C6C4;
                padding: 2px;
            }

            QMenu::item {
                padding: 6px 24px;
                border-radius: 2px;
            }

            QMenu::item:selected {
                background-color: #F3F2F1;
            }

            QMenu::item:hover {
                background-color: #F3F2F1;
            }

            QToolBar {
                background-color: #FFFFFF;
                border: none;
                border-bottom: 1px solid #D2D0CE;
                spacing: 2px;
                padding: 4px;
            }

            QToolButton {
                background-color: transparent;
                border: 1px solid transparent;
                border-radius: 2px;
                padding: 4px;
                color: #323130;
            }

            QToolButton:hover {
                background-color: #F3F2F1;
                border: 1px solid #EDEBE9;
            }

            QToolButton:pressed {
                background-color: #EDEBE9;
                border: 1px solid #E1DFDD;
            }

            QToolButton:checked {
                background-color: #EDEBE9;
                border: 1px solid #C8C6C4;
            }

            QStatusBar {
                background-color: #F3F2F1;
                color: #323130;
                border-top: 1px solid #D2D0CE;
                padding: 2px;
            }

            QStatusBar QLabel {
                color: #605E5C;
                padding: 2px 8px;
            }

            QTextEdit {
                background-color: #F3F2F1;
                color: #000000;
                border: none;
                selection-background-color: #0078D4;
                selection-color: #FFFFFF;
            }

            QDockWidget {
                background-color: #FAF9F8;
                border: 1px solid #EDEBE9;
                titlebar-close-icon: url(close.png);
                titlebar-normal-icon: url(float.png);
            }

            QDockWidget::title {
                background-color: #F3F2F1;
                padding: 6px;
                color: #323130;
                border-bottom: 1px solid #EDEBE9;
            }

            QPushButton {
                background-color: #FFFFFF;
                color: #323130;
                border: 1px solid #8A8886;
                border-radius: 2px;
                padding: 6px 16px;
                min-width: 64px;
                font-size: 14px;
            }

            QPushButton:hover {
                background-color: #F3F2F1;
                border: 1px solid #323130;
            }

            QPushButton:pressed {
                background-color: #EDEBE9;
                border: 1px solid #323130;
            }

            QPushButton:disabled {
                background-color: #F3F2F1;
                color: #A19F9D;
                border: 1px solid #EDEBE9;
            }

            QPushButton[primary="true"] {
                background-color: #0078D4;
                color: #FFFFFF;
                border: 1px solid #0078D4;
            }

            QPushButton[primary="true"]:hover {
                background-color: #106EBE;
                border: 1px solid #106EBE;
            }

            QTabWidget::pane {
                border: 1px solid #EDEBE9;
                background-color: #FFFFFF;
                border-top: none;
            }

            QTabBar::tab {
                background-color: #F3F2F1;
                color: #323130;
                border: 1px solid #EDEBE9;
                border-bottom: none;
                padding: 8px 16px;
                margin-right: 1px;
                border-top-left-radius: 2px;
                border-top-right-radius: 2px;
            }

            QTabBar::tab:selected {
                background-color: #FFFFFF;
                border-bottom: 2px solid #0078D4;
            }

            QTabBar::tab:hover:!selected {
                background-color: #E1DFDD;
            }

            QScrollBar:vertical {
                background-color: #FAF9F8;
                width: 14px;
                margin: 0px;
                border: none;
            }

            QScrollBar::handle:vertical {
                background-color: #C8C6C4;
                min-height: 30px;
                border-radius: 7px;
                margin: 2px;
            }

            QScrollBar::handle:vertical:hover {
                background-color: #A19F9D;
            }

            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }

            QScrollBar:horizontal {
                background-color: #FAF9F8;
                height: 14px;
                margin: 0px;
                border: none;
            }

            QScrollBar::handle:horizontal {
                background-color: #C8C6C4;
                min-width: 30px;
                border-radius: 7px;
                margin: 2px;
            }

            QScrollBar::handle:horizontal:hover {
                background-color: #A19F9D;
            }

            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
            }

            QListWidget, QTreeWidget, QTableWidget {
                background-color: #FFFFFF;
                color: #323130;
                border: 1px solid #EDEBE9;
                alternate-background-color: #FAF9F8;
                outline: none;
            }

            QListWidget::item:selected, QTreeWidget::item:selected, QTableWidget::item:selected {
                background-color: #EDEBE9;
                color: #323130;
            }

            QListWidget::item:hover, QTreeWidget::item:hover, QTableWidget::item:hover {
                background-color: #F3F2F1;
            }

            QLineEdit, QSpinBox, QDoubleSpinBox {
                background-color: #FFFFFF;
                color: #323130;
                border: 1px solid #8A8886;
                border-radius: 2px;
                padding: 6px 8px;
                selection-background-color: #0078D4;
                selection-color: #FFFFFF;
            }

            QLineEdit:hover, QSpinBox:hover, QDoubleSpinBox:hover {
                border: 1px solid #323130;
            }

            QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus {
                border: 2px solid #0078D4;
                padding: 5px 7px;
            }

            QComboBox {
                background-color: #FFFFFF;
                color: #323130;
                border: 1px solid #8A8886;
                border-radius: 2px;
                padding: 6px 8px;
                min-width: 100px;
            }

            QComboBox:hover {
                border: 1px solid #323130;
            }

            QComboBox:focus {
                border: 2px solid #0078D4;
            }

            QComboBox::drop-down {
                border: none;
                width: 20px;
            }

            QComboBox::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid #605E5C;
                margin-right: 4px;
            }

            QComboBox QAbstractItemView {
                background-color: #FFFFFF;
                color: #323130;
                border: 1px solid #C8C6C4;
                selection-background-color: #F3F2F1;
                selection-color: #323130;
                outline: none;
            }

            QFontComboBox {
                background-color: #FFFFFF;
                color: #323130;
                border: 1px solid #8A8886;
                border-radius: 2px;
                padding: 4px 6px;
            }
        """
        app.setStyleSheet(stylesheet)

    def apply_dark_theme(self, app: QApplication):
        """Apply dark theme to the application."""
        # Dark color palette
        dark_palette = QPalette()

        # Window colors
        dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.WindowText, QColor(255, 255, 255))

        # Base colors
        dark_palette.setColor(QPalette.Base, QColor(35, 35, 35))
        dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))

        # Text colors
        dark_palette.setColor(QPalette.Text, QColor(255, 255, 255))
        dark_palette.setColor(QPalette.PlaceholderText, QColor(128, 128, 128))

        # Button colors
        dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ButtonText, QColor(255, 255, 255))

        # Highlight colors
        dark_palette.setColor(QPalette.Highlight, QColor(0, 120, 215))
        dark_palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))

        # Link colors
        dark_palette.setColor(QPalette.Link, QColor(42, 130, 218))
        dark_palette.setColor(QPalette.LinkVisited, QColor(120, 100, 200))

        # ToolTip colors
        dark_palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
        dark_palette.setColor(QPalette.ToolTipText, QColor(0, 0, 0))

        # Disabled colors
        dark_palette.setColor(QPalette.Disabled, QPalette.WindowText, QColor(127, 127, 127))
        dark_palette.setColor(QPalette.Disabled, QPalette.Text, QColor(127, 127, 127))
        dark_palette.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(127, 127, 127))

        app.setPalette(dark_palette)

        # Dark theme stylesheet
        dark_stylesheet = """
            QMainWindow {
                background-color: #353535;
            }

            QMenuBar {
                background-color: #2d2d2d;
                color: #ffffff;
                border-bottom: 1px solid #1e1e1e;
            }

            QMenuBar::item {
                background-color: transparent;
                padding: 4px 10px;
            }

            QMenuBar::item:selected {
                background-color: #404040;
            }

            QMenu {
                background-color: #2d2d2d;
                color: #ffffff;
                border: 1px solid #1e1e1e;
            }

            QMenu::item:selected {
                background-color: #0078d7;
                color: #ffffff;
            }

            QToolBar {
                background-color: #2d2d2d;
                border: 1px solid #1e1e1e;
                spacing: 3px;
                padding: 2px;
            }

            QToolButton {
                background-color: transparent;
                border: 1px solid transparent;
                border-radius: 3px;
                padding: 3px;
                color: #ffffff;
            }

            QToolButton:hover {
                background-color: rgba(0, 120, 215, 0.2);
                border: 1px solid rgba(0, 120, 215, 0.5);
            }

            QToolButton:pressed {
                background-color: rgba(0, 120, 215, 0.3);
            }

            QStatusBar {
                background-color: #2d2d2d;
                color: #ffffff;
                border-top: 1px solid #1e1e1e;
            }

            QTextEdit {
                background-color: #1e1e1e;
                color: #ffffff;
                border: 1px solid #3e3e3e;
            }

            QDockWidget {
                background-color: #2d2d2d;
                color: #ffffff;
            }

            QDockWidget::title {
                background-color: #353535;
                padding: 4px;
                color: #ffffff;
            }

            QPushButton {
                background-color: #404040;
                color: #ffffff;
                border: 1px solid #3e3e3e;
                border-radius: 3px;
                padding: 5px 15px;
                min-width: 60px;
            }

            QPushButton:hover {
                background-color: #0078d7;
                border: 1px solid #005a9e;
            }

            QPushButton:pressed {
                background-color: #005a9e;
            }

            QPushButton:disabled {
                background-color: #404040;
                color: #808080;
            }

            QTabWidget::pane {
                border: 1px solid #3e3e3e;
                background-color: #2d2d2d;
            }

            QTabBar::tab {
                background-color: #353535;
                color: #ffffff;
                border: 1px solid #3e3e3e;
                padding: 5px 10px;
                margin-right: 2px;
            }

            QTabBar::tab:selected {
                background-color: #2d2d2d;
                border-bottom: 1px solid #2d2d2d;
            }

            QTabBar::tab:hover {
                background-color: #0078d7;
            }

            QScrollBar:vertical {
                background-color: #2d2d2d;
                width: 12px;
                margin: 0px;
            }

            QScrollBar::handle:vertical {
                background-color: #606060;
                min-height: 20px;
                border-radius: 6px;
                margin: 2px;
            }

            QScrollBar::handle:vertical:hover {
                background-color: #808080;
            }

            QScrollBar:horizontal {
                background-color: #2d2d2d;
                height: 12px;
                margin: 0px;
            }

            QScrollBar::handle:horizontal {
                background-color: #606060;
                min-width: 20px;
                border-radius: 6px;
                margin: 2px;
            }

            QScrollBar::handle:horizontal:hover {
                background-color: #808080;
            }

            QListWidget, QTreeWidget, QTableWidget {
                background-color: #1e1e1e;
                color: #ffffff;
                border: 1px solid #3e3e3e;
                alternate-background-color: #252525;
            }

            QListWidget::item:selected, QTreeWidget::item:selected, QTableWidget::item:selected {
                background-color: #0078d7;
                color: #ffffff;
            }

            QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox, QComboBox {
                background-color: #1e1e1e;
                color: #ffffff;
                border: 1px solid #3e3e3e;
                border-radius: 2px;
                padding: 3px;
            }

            QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {
                border: 1px solid #0078d7;
            }

            QComboBox::drop-down {
                border: none;
                background-color: #353535;
            }

            QComboBox QAbstractItemView {
                background-color: #2d2d2d;
                color: #ffffff;
                selection-background-color: #0078d7;
            }
        """
        app.setStyleSheet(dark_stylesheet)

    def get_current_theme(self) -> Theme:
        """Get the current theme."""
        return self.current_theme

    def toggle_theme(self):
        """Toggle between light and dark themes."""
        if self.current_theme == Theme.LIGHT:
            self.set_theme(Theme.DARK)
        else:
            self.set_theme(Theme.LIGHT)
