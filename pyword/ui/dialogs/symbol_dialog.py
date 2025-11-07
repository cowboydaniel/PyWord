"""Symbol Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QTableWidget, QTableWidgetItem,
                             QComboBox, QGroupBox, QFormLayout, QHeaderView)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from .base_dialog import BaseDialog


class SymbolDialog(BaseDialog):
    """Dialog for inserting special characters and symbols."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Symbol")
        self.selected_symbol = ""
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Font selection
        font_group = QGroupBox("Font")
        font_layout = QFormLayout()

        self.font_combo = QComboBox()
        self.font_combo.addItems([
            "Arial",
            "Times New Roman",
            "Symbol",
            "Wingdings",
            "Webdings",
        ])
        self.font_combo.currentTextChanged.connect(self.on_font_changed)
        font_layout.addRow("Font:", self.font_combo)

        # Character subset
        self.subset_combo = QComboBox()
        self.subset_combo.addItems([
            "Basic Latin",
            "Latin-1 Supplement",
            "Greek",
            "Arrows",
            "Mathematical Operators",
            "Geometric Shapes",
            "Miscellaneous Symbols",
        ])
        self.subset_combo.currentTextChanged.connect(self.on_subset_changed)
        font_layout.addRow("Subset:", self.subset_combo)

        font_group.setLayout(font_layout)
        layout.addWidget(font_group)

        # Symbol table
        symbols_group = QGroupBox("Characters")
        symbols_layout = QVBoxLayout()

        self.symbol_table = QTableWidget()
        self.symbol_table.setColumnCount(16)
        self.symbol_table.setRowCount(16)
        self.symbol_table.horizontalHeader().setVisible(False)
        self.symbol_table.verticalHeader().setVisible(False)
        self.symbol_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.symbol_table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.symbol_table.setSelectionMode(QTableWidget.SingleSelection)
        self.symbol_table.itemClicked.connect(self.on_symbol_clicked)
        self.symbol_table.itemDoubleClicked.connect(self.on_symbol_double_clicked)

        symbols_layout.addWidget(self.symbol_table)
        symbols_group.setLayout(symbols_layout)
        layout.addWidget(symbols_group)

        # Selected character info
        info_layout = QHBoxLayout()
        info_layout.addWidget(QLabel("Selected character:"))

        self.selected_label = QLabel("")
        self.selected_label.setFont(QFont("Arial", 24))
        self.selected_label.setMinimumWidth(60)
        self.selected_label.setAlignment(Qt.AlignCenter)
        self.selected_label.setFrameStyle(QLabel.Box | QLabel.Plain)
        info_layout.addWidget(self.selected_label)

        self.unicode_label = QLabel("")
        info_layout.addWidget(self.unicode_label)
        info_layout.addStretch()

        layout.addLayout(info_layout)

        # Recently used symbols (placeholder)
        recent_layout = QHBoxLayout()
        recent_layout.addWidget(QLabel("Recently used:"))
        recent_layout.addStretch()
        layout.addLayout(recent_layout)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.insert_button = QPushButton("Insert")
        self.insert_button.clicked.connect(self.accept)
        self.insert_button.setEnabled(False)
        button_layout.addWidget(self.insert_button)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.resize(600, 500)

        # Populate initial symbols
        self.populate_symbols()

    def populate_symbols(self):
        """Populate the symbol table with characters."""
        subset = self.subset_combo.currentText()

        # Define character ranges for different subsets
        ranges = {
            "Basic Latin": (0x0020, 0x007F),
            "Latin-1 Supplement": (0x00A0, 0x00FF),
            "Greek": (0x0370, 0x03FF),
            "Arrows": (0x2190, 0x21FF),
            "Mathematical Operators": (0x2200, 0x22FF),
            "Geometric Shapes": (0x25A0, 0x25FF),
            "Miscellaneous Symbols": (0x2600, 0x26FF),
        }

        start, end = ranges.get(subset, (0x0020, 0x007F))

        # Fill the table
        char_index = start
        font = QFont(self.font_combo.currentText(), 12)

        for row in range(16):
            for col in range(16):
                if char_index <= end:
                    try:
                        char = chr(char_index)
                        item = QTableWidgetItem(char)
                        item.setFont(font)
                        item.setTextAlignment(Qt.AlignCenter)
                        item.setData(Qt.UserRole, char_index)
                        self.symbol_table.setItem(row, col, item)
                        char_index += 1
                    except ValueError:
                        char_index += 1
                else:
                    break

    def on_font_changed(self, font_name):
        """Handle font selection change."""
        font = QFont(font_name, 12)
        for row in range(self.symbol_table.rowCount()):
            for col in range(self.symbol_table.columnCount()):
                item = self.symbol_table.item(row, col)
                if item:
                    item.setFont(font)

    def on_subset_changed(self, subset):
        """Handle character subset change."""
        self.populate_symbols()

    def on_symbol_clicked(self, item):
        """Handle symbol selection."""
        char_code = item.data(Qt.UserRole)
        if char_code:
            self.selected_symbol = chr(char_code)
            self.selected_label.setText(self.selected_symbol)
            self.unicode_label.setText(f"U+{char_code:04X}")
            self.insert_button.setEnabled(True)

    def on_symbol_double_clicked(self, item):
        """Handle symbol double-click (insert immediately)."""
        self.on_symbol_clicked(item)
        self.accept()

    def get_selected_symbol(self):
        """Get the selected symbol."""
        return self.selected_symbol
