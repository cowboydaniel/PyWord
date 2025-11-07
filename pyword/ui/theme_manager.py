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
        """Apply light theme to the application."""
        # Reset to default palette
        app.setStyleSheet("")
        app.setPalette(app.style().standardPalette())

        # Custom light theme stylesheet
        stylesheet = """
            QMainWindow {
                background-color: #ffffff;
            }

            QMenuBar {
                background-color: #f0f0f0;
                color: #000000;
                border-bottom: 1px solid #d0d0d0;
            }

            QMenuBar::item {
                background-color: transparent;
                padding: 4px 10px;
            }

            QMenuBar::item:selected {
                background-color: #e0e0e0;
            }

            QMenu {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #d0d0d0;
            }

            QMenu::item:selected {
                background-color: #0078d7;
                color: #ffffff;
            }

            QToolBar {
                background-color: #f3f3f3;
                border: 1px solid #d0d0d0;
                spacing: 3px;
                padding: 2px;
            }

            QToolButton {
                background-color: transparent;
                border: 1px solid transparent;
                border-radius: 3px;
                padding: 3px;
            }

            QToolButton:hover {
                background-color: rgba(0, 120, 215, 0.1);
                border: 1px solid rgba(0, 120, 215, 0.3);
            }

            QToolButton:pressed {
                background-color: rgba(0, 120, 215, 0.2);
            }

            QStatusBar {
                background-color: #f0f0f0;
                color: #000000;
                border-top: 1px solid #d0d0d0;
            }

            QTextEdit {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #d0d0d0;
            }

            QDockWidget {
                background-color: #f0f0f0;
                titlebar-close-icon: url(close.png);
                titlebar-normal-icon: url(float.png);
            }

            QDockWidget::title {
                background-color: #e0e0e0;
                padding: 4px;
            }

            QPushButton {
                background-color: #f0f0f0;
                color: #000000;
                border: 1px solid #d0d0d0;
                border-radius: 3px;
                padding: 5px 15px;
                min-width: 60px;
            }

            QPushButton:hover {
                background-color: #e5f3ff;
                border: 1px solid #0078d7;
            }

            QPushButton:pressed {
                background-color: #cce4f7;
            }

            QPushButton:disabled {
                background-color: #f0f0f0;
                color: #808080;
            }

            QTabWidget::pane {
                border: 1px solid #d0d0d0;
                background-color: #ffffff;
            }

            QTabBar::tab {
                background-color: #f0f0f0;
                color: #000000;
                border: 1px solid #d0d0d0;
                padding: 5px 10px;
                margin-right: 2px;
            }

            QTabBar::tab:selected {
                background-color: #ffffff;
                border-bottom: 1px solid #ffffff;
            }

            QTabBar::tab:hover {
                background-color: #e5f3ff;
            }

            QScrollBar:vertical {
                background-color: #f0f0f0;
                width: 12px;
                margin: 0px;
            }

            QScrollBar::handle:vertical {
                background-color: #c0c0c0;
                min-height: 20px;
                border-radius: 6px;
                margin: 2px;
            }

            QScrollBar::handle:vertical:hover {
                background-color: #a0a0a0;
            }

            QScrollBar:horizontal {
                background-color: #f0f0f0;
                height: 12px;
                margin: 0px;
            }

            QScrollBar::handle:horizontal {
                background-color: #c0c0c0;
                min-width: 20px;
                border-radius: 6px;
                margin: 2px;
            }

            QScrollBar::handle:horizontal:hover {
                background-color: #a0a0a0;
            }

            QListWidget, QTreeWidget, QTableWidget {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #d0d0d0;
                alternate-background-color: #f8f8f8;
            }

            QListWidget::item:selected, QTreeWidget::item:selected, QTableWidget::item:selected {
                background-color: #0078d7;
                color: #ffffff;
            }

            QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox, QComboBox {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #d0d0d0;
                border-radius: 2px;
                padding: 3px;
            }

            QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {
                border: 1px solid #0078d7;
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
