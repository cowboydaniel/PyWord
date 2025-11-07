"""About Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QTextBrowser)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap, QIcon
from .base_dialog import BaseDialog


class AboutDialog(BaseDialog):
    """About dialog showing application information."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About PyWord")
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Logo/Icon
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)
        icon = QIcon.fromTheme("accessories-text-editor")
        if not icon.isNull():
            pixmap = icon.pixmap(64, 64)
            logo_label.setPixmap(pixmap)
        layout.addWidget(logo_label)

        # Application name and version
        name_label = QLabel("<h1>PyWord</h1>")
        name_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(name_label)

        version_label = QLabel("<p>Version 1.0.0</p>")
        version_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(version_label)

        # Description
        description = QTextBrowser()
        description.setOpenExternalLinks(True)
        description.setHtml("""
        <p style="text-align: center;">
        <b>A Python-based Word Processor</b><br><br>

        PyWord is a modern, feature-rich word processor built with Python and PySide6.<br><br>

        Features include:<br>
        • Document editing with rich text formatting<br>
        • Tables, images, and hyperlinks<br>
        • Headers and footers<br>
        • Styles and themes<br>
        • Page layout and printing<br>
        • And much more!<br><br>

        <b>Copyright © 2024</b><br>
        Licensed under the MIT License<br><br>

        Built with PySide6 (Qt for Python)
        </p>
        """)
        description.setMaximumHeight(300)
        layout.addWidget(description)

        # Close button
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.accept)
        button_layout.addWidget(self.close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.resize(400, 500)
