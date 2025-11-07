"""
Advanced Table Formatting for PyWord.

This module extends the basic table functionality with advanced formatting options
including cell merging, splitting, borders, shading, alignment, and table styles.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
                               QPushButton, QGroupBox, QSpinBox, QDoubleSpinBox,
                               QColorDialog, QCheckBox, QTabWidget, QWidget,
                               QFormLayout, QLineEdit, QSlider, QRadioButton, QButtonGroup)
from PySide6.QtCore import Qt
from PySide6.QtGui import (QColor, QTextTableFormat, QTextTableCellFormat, QTextLength,
                          QTextFrameFormat, QTextCharFormat, QBrush, QPen)


class TableStyle:
    """Predefined table styles."""

    PLAIN = "Plain"
    GRID = "Grid"
    MODERN = "Modern"
    ELEGANT = "Elegant"
    PROFESSIONAL = "Professional"
    COLORFUL = "Colorful"

    @staticmethod
    def get_style_format(style_name):
        """Get table format for a predefined style."""
        table_format = QTextTableFormat()

        if style_name == TableStyle.PLAIN:
            table_format.setBorder(1)
            table_format.setBorderBrush(QBrush(QColor(Qt.GlobalColor.black)))
            table_format.setCellPadding(5)
            table_format.setCellSpacing(0)

        elif style_name == TableStyle.GRID:
            table_format.setBorder(2)
            table_format.setBorderBrush(QBrush(QColor(Qt.GlobalColor.darkGray)))
            table_format.setCellPadding(8)
            table_format.setCellSpacing(0)

        elif style_name == TableStyle.MODERN:
            table_format.setBorder(0)
            table_format.setBorderBrush(QBrush(QColor(70, 130, 180)))
            table_format.setCellPadding(10)
            table_format.setCellSpacing(0)

        elif style_name == TableStyle.ELEGANT:
            table_format.setBorder(1)
            table_format.setBorderBrush(QBrush(QColor(128, 128, 128)))
            table_format.setCellPadding(12)
            table_format.setCellSpacing(0)

        elif style_name == TableStyle.PROFESSIONAL:
            table_format.setBorder(1)
            table_format.setBorderBrush(QBrush(QColor(0, 0, 0)))
            table_format.setCellPadding(8)
            table_format.setCellSpacing(0)

        elif style_name == TableStyle.COLORFUL:
            table_format.setBorder(2)
            table_format.setBorderBrush(QBrush(QColor(255, 165, 0)))
            table_format.setCellPadding(10)
            table_format.setCellSpacing(2)

        return table_format


class AdvancedTableManager:
    """Manages advanced table formatting operations."""

    def __init__(self, parent):
        self.parent = parent
        self.current_table = None

    def get_current_table(self):
        """Get the table at the current cursor position."""
        cursor = self.parent.textCursor()
        return cursor.currentTable()

    def merge_cells(self, start_row, start_col, num_rows, num_cols):
        """Merge a range of cells."""
        table = self.get_current_table()
        if table:
            table.mergeCells(start_row, start_col, num_rows, num_cols)
            return True
        return False

    def split_cell(self, row, col, num_rows, num_cols):
        """Split a cell into multiple cells."""
        table = self.get_current_table()
        if table:
            table.splitCell(row, col, num_rows, num_cols)
            return True
        return False

    def set_cell_background(self, row, col, color):
        """Set the background color of a cell."""
        table = self.get_current_table()
        if table:
            cell = table.cellAt(row, col)
            if cell.isValid():
                cell_format = cell.format().toTableCellFormat()
                cell_format.setBackground(QBrush(color))
                cell.setFormat(cell_format)
                return True
        return False

    def set_cell_border(self, row, col, width, color, style="solid"):
        """Set the border of a cell."""
        table = self.get_current_table()
        if table:
            cell = table.cellAt(row, col)
            if cell.isValid():
                cell_format = cell.format().toTableCellFormat()
                cell_format.setBorder(width)
                cell_format.setBorderBrush(QBrush(color))
                cell.setFormat(cell_format)
                return True
        return False

    def set_cell_padding(self, row, col, padding):
        """Set the padding of a cell."""
        table = self.get_current_table()
        if table:
            cell = table.cellAt(row, col)
            if cell.isValid():
                cell_format = cell.format().toTableCellFormat()
                cell_format.setTopPadding(padding)
                cell_format.setBottomPadding(padding)
                cell_format.setLeftPadding(padding)
                cell_format.setRightPadding(padding)
                cell.setFormat(cell_format)
                return True
        return False

    def set_cell_alignment(self, row, col, alignment):
        """Set the text alignment within a cell."""
        table = self.get_current_table()
        if table:
            cell = table.cellAt(row, col)
            if cell.isValid():
                cursor = cell.firstCursorPosition()
                block_format = cursor.blockFormat()
                block_format.setAlignment(alignment)
                cursor.setBlockFormat(block_format)
                return True
        return False

    def apply_table_style(self, style_name):
        """Apply a predefined style to the current table."""
        table = self.get_current_table()
        if table:
            table_format = TableStyle.get_style_format(style_name)
            table.setFormat(table_format)

            # Apply header row formatting for certain styles
            if style_name in [TableStyle.MODERN, TableStyle.PROFESSIONAL, TableStyle.ELEGANT]:
                self.format_header_row(table)

            return True
        return False

    def format_header_row(self, table):
        """Apply special formatting to the header row."""
        if table.rows() < 1:
            return

        for col in range(table.columns()):
            cell = table.cellAt(0, col)
            if cell.isValid():
                # Set cell background
                cell_format = cell.format().toTableCellFormat()
                cell_format.setBackground(QBrush(QColor(70, 130, 180)))
                cell.setFormat(cell_format)

                # Set text formatting
                cursor = cell.firstCursorPosition()
                char_format = cursor.charFormat()
                char_format.setFontWeight(700)  # Bold
                char_format.setForeground(QBrush(QColor(Qt.GlobalColor.white)))
                cursor.select(cursor.SelectionType.Document)
                cursor.mergeCharFormat(char_format)

    def set_table_width(self, width, unit="percent"):
        """Set the width of the current table."""
        table = self.get_current_table()
        if table:
            table_format = table.format()

            if unit == "percent":
                table_format.setWidth(QTextLength(QTextLength.Type.PercentageLength, width))
            else:  # pixels
                table_format.setWidth(QTextLength(QTextLength.Type.FixedLength, width))

            table.setFormat(table_format)
            return True
        return False

    def set_table_alignment(self, alignment):
        """Set the alignment of the table within the document."""
        table = self.get_current_table()
        if table:
            table_format = table.format()
            table_format.setAlignment(alignment)
            table.setFormat(table_format)
            return True
        return False

    def set_column_width(self, col, width):
        """Set the width of a specific column."""
        table = self.get_current_table()
        if table and 0 <= col < table.columns():
            table_format = table.format()
            constraints = table_format.columnWidthConstraints()

            # Modify or add constraint for the column
            if col < len(constraints):
                constraints[col] = QTextLength(QTextLength.Type.FixedLength, width)
            else:
                while len(constraints) <= col:
                    constraints.append(QTextLength(QTextLength.Type.VariableLength, 0))
                constraints[col] = QTextLength(QTextLength.Type.FixedLength, width)

            table_format.setColumnWidthConstraints(constraints)
            table.setFormat(table_format)
            return True
        return False

    def distribute_columns_evenly(self):
        """Distribute column widths evenly."""
        table = self.get_current_table()
        if table:
            table_format = table.format()
            num_cols = table.columns()

            # Create equal width constraints
            constraints = []
            for i in range(num_cols):
                constraints.append(QTextLength(QTextLength.Type.PercentageLength, 100.0 / num_cols))

            table_format.setColumnWidthConstraints(constraints)
            table.setFormat(table_format)
            return True
        return False

    def auto_fit_contents(self):
        """Auto-fit table to contents."""
        table = self.get_current_table()
        if table:
            table_format = table.format()
            table_format.setWidth(QTextLength(QTextLength.Type.VariableLength, 0))
            table.setFormat(table_format)
            return True
        return False


class AdvancedTableDialog(QDialog):
    """Dialog for advanced table formatting."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Advanced Table Formatting")
        self.setModal(True)
        self.setMinimumWidth(600)

        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Tab widget
        tabs = QTabWidget()

        # Style tab
        style_tab = self.create_style_tab()
        tabs.addTab(style_tab, "Table Style")

        # Cell formatting tab
        cell_tab = self.create_cell_tab()
        tabs.addTab(cell_tab, "Cell Formatting")

        # Borders tab
        borders_tab = self.create_borders_tab()
        tabs.addTab(borders_tab, "Borders")

        # Size tab
        size_tab = self.create_size_tab()
        tabs.addTab(size_tab, "Size & Alignment")

        layout.addWidget(tabs)

        # Buttons
        button_layout = QHBoxLayout()
        apply_button = QPushButton("Apply")
        apply_button.clicked.connect(self.apply_formatting)
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(apply_button)
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def create_style_tab(self):
        """Create the table style tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        style_group = QGroupBox("Predefined Styles")
        style_layout = QVBoxLayout()

        self.style_combo = QComboBox()
        self.style_combo.addItems([
            TableStyle.PLAIN,
            TableStyle.GRID,
            TableStyle.MODERN,
            TableStyle.ELEGANT,
            TableStyle.PROFESSIONAL,
            TableStyle.COLORFUL
        ])

        style_layout.addWidget(QLabel("Choose a table style:"))
        style_layout.addWidget(self.style_combo)

        apply_style_button = QPushButton("Apply Style")
        apply_style_button.clicked.connect(self.apply_style)
        style_layout.addWidget(apply_style_button)

        style_group.setLayout(style_layout)
        layout.addWidget(style_group)

        # Header row options
        header_group = QGroupBox("Header Row")
        header_layout = QVBoxLayout()

        self.format_header_checkbox = QCheckBox("Format first row as header")
        self.format_header_checkbox.setChecked(True)
        header_layout.addWidget(self.format_header_checkbox)

        format_header_button = QPushButton("Apply Header Formatting")
        format_header_button.clicked.connect(self.format_header)
        header_layout.addWidget(format_header_button)

        header_group.setLayout(header_layout)
        layout.addWidget(header_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_cell_tab(self):
        """Create the cell formatting tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Cell selection
        selection_group = QGroupBox("Cell Selection")
        selection_layout = QFormLayout()

        self.cell_row_spin = QSpinBox()
        self.cell_row_spin.setMinimum(0)
        self.cell_col_spin = QSpinBox()
        self.cell_col_spin.setMinimum(0)

        selection_layout.addRow("Row:", self.cell_row_spin)
        selection_layout.addRow("Column:", self.cell_col_spin)

        selection_group.setLayout(selection_layout)
        layout.addWidget(selection_group)

        # Cell background
        bg_group = QGroupBox("Cell Background")
        bg_layout = QHBoxLayout()

        self.cell_bg_button = QPushButton("Choose Color")
        self.cell_bg_button.clicked.connect(self.choose_cell_background)
        bg_layout.addWidget(self.cell_bg_button)

        bg_group.setLayout(bg_layout)
        layout.addWidget(bg_group)

        # Cell alignment
        align_group = QGroupBox("Cell Alignment")
        align_layout = QVBoxLayout()

        self.align_button_group = QButtonGroup()
        self.align_left = QRadioButton("Left")
        self.align_center = QRadioButton("Center")
        self.align_right = QRadioButton("Right")
        self.align_justify = QRadioButton("Justify")

        self.align_button_group.addButton(self.align_left)
        self.align_button_group.addButton(self.align_center)
        self.align_button_group.addButton(self.align_right)
        self.align_button_group.addButton(self.align_justify)

        self.align_left.setChecked(True)

        align_layout.addWidget(self.align_left)
        align_layout.addWidget(self.align_center)
        align_layout.addWidget(self.align_right)
        align_layout.addWidget(self.align_justify)

        apply_align_button = QPushButton("Apply Alignment")
        apply_align_button.clicked.connect(self.apply_cell_alignment)
        align_layout.addWidget(apply_align_button)

        align_group.setLayout(align_layout)
        layout.addWidget(align_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_borders_tab(self):
        """Create the borders tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        border_group = QGroupBox("Border Settings")
        border_layout = QFormLayout()

        self.border_width_spin = QDoubleSpinBox()
        self.border_width_spin.setMinimum(0)
        self.border_width_spin.setMaximum(10)
        self.border_width_spin.setValue(1)
        self.border_width_spin.setSingleStep(0.5)

        self.border_color_button = QPushButton("Choose Color")
        self.border_color_button.clicked.connect(self.choose_border_color)

        border_layout.addRow("Border Width:", self.border_width_spin)
        border_layout.addRow("Border Color:", self.border_color_button)

        border_group.setLayout(border_layout)
        layout.addWidget(border_group)

        # Cell padding
        padding_group = QGroupBox("Cell Padding")
        padding_layout = QFormLayout()

        self.cell_padding_spin = QDoubleSpinBox()
        self.cell_padding_spin.setMinimum(0)
        self.cell_padding_spin.setMaximum(50)
        self.cell_padding_spin.setValue(5)

        padding_layout.addRow("Padding:", self.cell_padding_spin)

        padding_group.setLayout(padding_layout)
        layout.addWidget(padding_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_size_tab(self):
        """Create the size and alignment tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Table width
        width_group = QGroupBox("Table Width")
        width_layout = QFormLayout()

        self.table_width_spin = QSpinBox()
        self.table_width_spin.setMinimum(1)
        self.table_width_spin.setMaximum(100)
        self.table_width_spin.setValue(100)

        self.width_unit_combo = QComboBox()
        self.width_unit_combo.addItems(["Percent", "Pixels"])

        width_layout.addRow("Width:", self.table_width_spin)
        width_layout.addRow("Unit:", self.width_unit_combo)

        apply_width_button = QPushButton("Apply Width")
        apply_width_button.clicked.connect(self.apply_table_width)
        width_layout.addRow(apply_width_button)

        width_group.setLayout(width_layout)
        layout.addWidget(width_group)

        # Table alignment
        table_align_group = QGroupBox("Table Alignment")
        table_align_layout = QVBoxLayout()

        self.table_align_button_group = QButtonGroup()
        self.table_align_left = QRadioButton("Left")
        self.table_align_center = QRadioButton("Center")
        self.table_align_right = QRadioButton("Right")

        self.table_align_button_group.addButton(self.table_align_left)
        self.table_align_button_group.addButton(self.table_align_center)
        self.table_align_button_group.addButton(self.table_align_right)

        self.table_align_left.setChecked(True)

        table_align_layout.addWidget(self.table_align_left)
        table_align_layout.addWidget(self.table_align_center)
        table_align_layout.addWidget(self.table_align_right)

        apply_table_align_button = QPushButton("Apply Alignment")
        apply_table_align_button.clicked.connect(self.apply_table_alignment)
        table_align_layout.addWidget(apply_table_align_button)

        table_align_group.setLayout(table_align_layout)
        layout.addWidget(table_align_group)

        # Quick actions
        quick_group = QGroupBox("Quick Actions")
        quick_layout = QVBoxLayout()

        distribute_button = QPushButton("Distribute Columns Evenly")
        distribute_button.clicked.connect(self.manager.distribute_columns_evenly)

        autofit_button = QPushButton("Auto-Fit to Contents")
        autofit_button.clicked.connect(self.manager.auto_fit_contents)

        quick_layout.addWidget(distribute_button)
        quick_layout.addWidget(autofit_button)

        quick_group.setLayout(quick_layout)
        layout.addWidget(quick_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def apply_style(self):
        """Apply the selected table style."""
        style = self.style_combo.currentText()
        self.manager.apply_table_style(style)

    def format_header(self):
        """Format the header row."""
        table = self.manager.get_current_table()
        if table and self.format_header_checkbox.isChecked():
            self.manager.format_header_row(table)

    def choose_cell_background(self):
        """Choose cell background color."""
        color = QColorDialog.getColor()
        if color.isValid():
            row = self.cell_row_spin.value()
            col = self.cell_col_spin.value()
            self.manager.set_cell_background(row, col, color)

    def choose_border_color(self):
        """Choose border color."""
        color = QColorDialog.getColor()
        if color.isValid():
            self.border_color = color

    def apply_cell_alignment(self):
        """Apply cell alignment."""
        row = self.cell_row_spin.value()
        col = self.cell_col_spin.value()

        if self.align_left.isChecked():
            alignment = Qt.AlignmentFlag.AlignLeft
        elif self.align_center.isChecked():
            alignment = Qt.AlignmentFlag.AlignCenter
        elif self.align_right.isChecked():
            alignment = Qt.AlignmentFlag.AlignRight
        else:
            alignment = Qt.AlignmentFlag.AlignJustify

        self.manager.set_cell_alignment(row, col, alignment)

    def apply_table_width(self):
        """Apply table width."""
        width = self.table_width_spin.value()
        unit = "percent" if self.width_unit_combo.currentText() == "Percent" else "pixels"
        self.manager.set_table_width(width, unit)

    def apply_table_alignment(self):
        """Apply table alignment."""
        if self.table_align_left.isChecked():
            alignment = Qt.AlignmentFlag.AlignLeft
        elif self.table_align_center.isChecked():
            alignment = Qt.AlignmentFlag.AlignCenter
        else:
            alignment = Qt.AlignmentFlag.AlignRight

        self.manager.set_table_alignment(alignment)

    def apply_formatting(self):
        """Apply all formatting settings."""
        self.apply_style()
        if self.format_header_checkbox.isChecked():
            self.format_header()

    def accept(self):
        """Handle OK button click."""
        self.apply_formatting()
        super().accept()
