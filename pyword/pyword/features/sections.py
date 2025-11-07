"""
Section and Section Break Management for PyWord.

This module handles sections and section breaks in documents, allowing
different parts of a document to have different formatting, headers/footers,
page orientation, and margins.
"""

from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QPushButton, QGroupBox
from PySide6.QtCore import Qt
from PySide6.QtGui import QTextFrameFormat, QTextCursor, QTextBlockFormat


class SectionBreakType:
    """Enumeration of section break types."""
    NEXT_PAGE = "Next Page"
    CONTINUOUS = "Continuous"
    EVEN_PAGE = "Even Page"
    ODD_PAGE = "Odd Page"


class Section:
    """Represents a document section with its own formatting properties."""

    def __init__(self):
        self.break_type = SectionBreakType.NEXT_PAGE
        self.start_position = 0
        self.end_position = -1
        self.header = None
        self.footer = None
        self.page_orientation = "Portrait"
        self.margins = {"top": 1.0, "bottom": 1.0, "left": 1.0, "right": 1.0}
        self.columns = 1
        self.page_numbering_start = 1
        self.different_first_page = False
        self.different_odd_even = False

    def apply_formatting(self, cursor):
        """Apply section-specific formatting to the cursor position."""
        block_format = QTextBlockFormat()

        # Apply section properties
        if self.break_type == SectionBreakType.NEXT_PAGE:
            block_format.setPageBreakPolicy(QTextBlockFormat.PageBreak_AlwaysBefore)

        cursor.setBlockFormat(block_format)


class SectionManager:
    """Manages sections and section breaks in a document."""

    def __init__(self, parent):
        self.parent = parent
        self.sections = []
        self.current_section_index = 0

        # Initialize with default section
        default_section = Section()
        self.sections.append(default_section)

    def insert_section_break(self, break_type=SectionBreakType.NEXT_PAGE):
        """Insert a section break at the current cursor position."""
        cursor = self.parent.textCursor()
        if cursor.isNull():
            return None

        # Create new section
        new_section = Section()
        new_section.break_type = break_type
        new_section.start_position = cursor.position()

        # Update previous section's end position
        if self.sections:
            if self.current_section_index < len(self.sections):
                self.sections[self.current_section_index].end_position = cursor.position()

        # Insert section break marker
        block_format = QTextBlockFormat()

        if break_type == SectionBreakType.NEXT_PAGE:
            block_format.setPageBreakPolicy(QTextBlockFormat.PageBreak_AlwaysBefore)
            cursor.insertBlock(block_format)
            cursor.insertText(f"\n═══ Section Break ({break_type}) ═══\n")
        elif break_type == SectionBreakType.CONTINUOUS:
            cursor.insertBlock(block_format)
            cursor.insertText(f"\n─── Section Break ({break_type}) ───\n")
        elif break_type == SectionBreakType.EVEN_PAGE:
            block_format.setPageBreakPolicy(QTextBlockFormat.PageBreak_AlwaysBefore)
            cursor.insertBlock(block_format)
            cursor.insertText(f"\n═══ Section Break ({break_type}) ═══\n")
        elif break_type == SectionBreakType.ODD_PAGE:
            block_format.setPageBreakPolicy(QTextBlockFormat.PageBreak_AlwaysBefore)
            cursor.insertBlock(block_format)
            cursor.insertText(f"\n═══ Section Break ({break_type}) ═══\n")

        # Add new section
        self.sections.append(new_section)
        self.current_section_index = len(self.sections) - 1

        return new_section

    def get_current_section(self):
        """Get the section at the current cursor position."""
        cursor = self.parent.textCursor()
        position = cursor.position()

        for i, section in enumerate(self.sections):
            if section.start_position <= position and (section.end_position == -1 or position <= section.end_position):
                self.current_section_index = i
                return section

        # Return last section if not found
        return self.sections[-1] if self.sections else None

    def get_section_by_index(self, index):
        """Get a section by its index."""
        if 0 <= index < len(self.sections):
            return self.sections[index]
        return None

    def delete_section_break(self):
        """Delete the section break at the current cursor position."""
        cursor = self.parent.textCursor()

        # Find and remove section break marker
        cursor.select(QTextCursor.SelectionType.LineUnderCursor)
        text = cursor.selectedText()

        if "Section Break" in text:
            cursor.removeSelectedText()

            # Remove section from list
            section = self.get_current_section()
            if section and section != self.sections[0]:  # Don't delete first section
                self.sections.remove(section)
                self.current_section_index = max(0, self.current_section_index - 1)

            return True
        return False

    def update_section_properties(self, section, **properties):
        """Update properties of a section."""
        for key, value in properties.items():
            if hasattr(section, key):
                setattr(section, key, value)

    def get_section_count(self):
        """Get the total number of sections."""
        return len(self.sections)

    def navigate_to_section(self, index):
        """Navigate to a specific section."""
        if 0 <= index < len(self.sections):
            section = self.sections[index]
            cursor = self.parent.textCursor()
            cursor.setPosition(section.start_position)
            self.parent.setTextCursor(cursor)
            self.current_section_index = index
            return True
        return False


class SectionBreakDialog(QDialog):
    """Dialog for inserting section breaks."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Section Break")
        self.setModal(True)
        self.selected_break_type = SectionBreakType.NEXT_PAGE

        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Section break type group
        type_group = QGroupBox("Section Break Type")
        type_layout = QVBoxLayout()

        type_label = QLabel("Choose where to start the new section:")
        type_layout.addWidget(type_label)

        self.break_type_combo = QComboBox()
        self.break_type_combo.addItem(SectionBreakType.NEXT_PAGE, SectionBreakType.NEXT_PAGE)
        self.break_type_combo.addItem(SectionBreakType.CONTINUOUS, SectionBreakType.CONTINUOUS)
        self.break_type_combo.addItem(SectionBreakType.EVEN_PAGE, SectionBreakType.EVEN_PAGE)
        self.break_type_combo.addItem(SectionBreakType.ODD_PAGE, SectionBreakType.ODD_PAGE)
        type_layout.addWidget(self.break_type_combo)

        # Descriptions
        description = QLabel(
            "<b>Next Page:</b> Start section on the next page<br>"
            "<b>Continuous:</b> Start section on the same page<br>"
            "<b>Even Page:</b> Start section on the next even-numbered page<br>"
            "<b>Odd Page:</b> Start section on the next odd-numbered page"
        )
        description.setWordWrap(True)
        type_layout.addWidget(description)

        type_group.setLayout(type_layout)
        layout.addWidget(type_group)

        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def accept(self):
        """Handle OK button click."""
        self.selected_break_type = self.break_type_combo.currentData()
        super().accept()

    def get_break_type(self):
        """Get the selected break type."""
        return self.selected_break_type
