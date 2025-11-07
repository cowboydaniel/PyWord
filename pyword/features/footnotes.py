"""
Footnotes and Endnotes System for PyWord.

This module provides comprehensive support for footnotes and endnotes,
including automatic numbering, positioning, and formatting.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit,
                               QPushButton, QListWidget, QListWidgetItem, QGroupBox,
                               QRadioButton, QComboBox, QSpinBox, QCheckBox, QMessageBox,
                               QInputDialog, QButtonGroup, QFormLayout, QWidget)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QTextCursor, QTextCharFormat, QFont, QColor, QTextBlockFormat
from datetime import datetime
import uuid


class Note:
    """Represents a footnote or endnote."""

    def __init__(self, text, note_type='footnote', position=0, reference_mark=None):
        self.id = str(uuid.uuid4())
        self.text = text
        self.note_type = note_type  # 'footnote' or 'endnote'
        self.position = position  # Position in document where referenced
        self.reference_mark = reference_mark  # Custom reference mark (e.g., *, †, 1, i)
        self.number = None  # Auto-assigned number
        self.created = datetime.now()
        self.modified = datetime.now()

    def to_dict(self):
        """Convert note to dictionary for serialization."""
        return {
            'id': self.id,
            'text': self.text,
            'note_type': self.note_type,
            'position': self.position,
            'reference_mark': self.reference_mark,
            'number': self.number,
            'created': self.created.isoformat(),
            'modified': self.modified.isoformat()
        }

    @staticmethod
    def from_dict(data):
        """Create Note from dictionary."""
        note = Note(
            data['text'],
            data.get('note_type', 'footnote'),
            data.get('position', 0),
            data.get('reference_mark')
        )
        note.id = data['id']
        note.number = data.get('number')
        if 'created' in data:
            note.created = datetime.fromisoformat(data['created'])
        if 'modified' in data:
            note.modified = datetime.fromisoformat(data['modified'])
        return note


class FootnotesManager:
    """Manages footnotes and endnotes in a document."""

    def __init__(self, parent):
        self.parent = parent  # Reference to text editor
        self.notes = []

        # Settings
        self.footnote_numbering = 'numeric'  # 'numeric', 'alphabetic', 'roman', 'symbols'
        self.endnote_numbering = 'numeric'
        self.footnote_start_number = 1
        self.endnote_start_number = 1
        self.restart_numbering = 'continuous'  # 'continuous', 'each_page', 'each_section'
        self.footnote_position = 'bottom_of_page'  # 'bottom_of_page', 'below_text'
        self.endnote_position = 'end_of_document'  # 'end_of_document', 'end_of_section'

        # Formatting
        self.footnote_separator = True
        self.separator_length = 50  # pixels

    def add_footnote(self, text, position=None, custom_mark=None):
        """Add a new footnote."""
        if position is None:
            cursor = self.parent.textCursor()
            position = cursor.position()

        note = Note(text, 'footnote', position, custom_mark)
        self.notes.append(note)
        self._renumber_notes()
        self._insert_reference_mark(note, position)

        return note

    def add_endnote(self, text, position=None, custom_mark=None):
        """Add a new endnote."""
        if position is None:
            cursor = self.parent.textCursor()
            position = cursor.position()

        note = Note(text, 'endnote', position, custom_mark)
        self.notes.append(note)
        self._renumber_notes()
        self._insert_reference_mark(note, position)

        return note

    def edit_note(self, note_id, new_text):
        """Edit an existing note."""
        note = self.get_note_by_id(note_id)
        if note:
            note.text = new_text
            note.modified = datetime.now()
            return True
        return False

    def delete_note(self, note_id):
        """Delete a note."""
        note = self.get_note_by_id(note_id)
        if not note:
            return False

        # Remove reference mark from document
        self._remove_reference_mark(note)

        # Remove note
        self.notes.remove(note)
        self._renumber_notes()

        return True

    def get_note_by_id(self, note_id):
        """Get a note by its ID."""
        for note in self.notes:
            if note.id == note_id:
                return note
        return None

    def get_notes_by_type(self, note_type):
        """Get all notes of a specific type."""
        return [n for n in self.notes if n.note_type == note_type]

    def get_footnotes(self):
        """Get all footnotes."""
        return self.get_notes_by_type('footnote')

    def get_endnotes(self):
        """Get all endnotes."""
        return self.get_notes_by_type('endnote')

    def _renumber_notes(self):
        """Renumber all notes based on their position in document."""
        # Sort notes by position
        footnotes = sorted(self.get_footnotes(), key=lambda n: n.position)
        endnotes = sorted(self.get_endnotes(), key=lambda n: n.position)

        # Renumber footnotes
        for i, note in enumerate(footnotes):
            if note.reference_mark is None:
                note.number = self._format_number(
                    i + self.footnote_start_number,
                    self.footnote_numbering
                )

        # Renumber endnotes
        for i, note in enumerate(endnotes):
            if note.reference_mark is None:
                note.number = self._format_number(
                    i + self.endnote_start_number,
                    self.endnote_numbering
                )

    def _format_number(self, number, style):
        """Format a number according to the specified style."""
        if style == 'numeric':
            return str(number)
        elif style == 'alphabetic':
            # Convert to lowercase letters (a, b, c, ..., z, aa, ab, ...)
            result = ""
            while number > 0:
                number -= 1
                result = chr(ord('a') + (number % 26)) + result
                number //= 26
            return result
        elif style == 'roman':
            # Convert to roman numerals
            return self._to_roman(number)
        elif style == 'symbols':
            # Use symbols (*, †, ‡, §, ¶, #)
            symbols = ['*', '†', '‡', '§', '¶', '#']
            if number <= len(symbols):
                return symbols[number - 1]
            else:
                # Repeat symbols for larger numbers
                symbol_index = (number - 1) % len(symbols)
                repeat_count = (number - 1) // len(symbols) + 1
                return symbols[symbol_index] * repeat_count

        return str(number)

    def _to_roman(self, number):
        """Convert number to roman numerals."""
        values = [
            (1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
            (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
            (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')
        ]

        result = ''
        for value, numeral in values:
            count = number // value
            if count:
                result += numeral * count
                number -= value * count

        return result.lower()

    def _insert_reference_mark(self, note, position):
        """Insert reference mark in the document."""
        cursor = QTextCursor(self.parent.document())
        cursor.setPosition(position)

        # Format for superscript reference mark
        char_format = QTextCharFormat()
        char_format.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSuperScript)
        char_format.setForeground(QColor(0, 0, 255))  # Blue
        char_format.setFontPointSize(8)

        # Get reference mark text
        mark_text = note.reference_mark if note.reference_mark else note.number

        # Insert mark
        cursor.setCharFormat(char_format)
        cursor.insertText(mark_text)

    def _remove_reference_mark(self, note):
        """Remove reference mark from document."""
        # This is simplified - in a real implementation,
        # you'd track the exact position of the mark
        cursor = QTextCursor(self.parent.document())
        cursor.setPosition(note.position)

        # Search for the reference mark
        mark_text = note.reference_mark if note.reference_mark else note.number
        if mark_text:
            # Move cursor to select the mark
            cursor.movePosition(
                QTextCursor.MoveOperation.Right,
                QTextCursor.MoveMode.KeepAnchor,
                len(str(mark_text))
            )
            cursor.removeSelectedText()

    def insert_footnotes_section(self):
        """Insert footnotes section at the bottom of the current page/document."""
        footnotes = self.get_footnotes()
        if not footnotes:
            return

        # Move to end of document (simplified - should be bottom of page)
        cursor = self.parent.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)

        # Insert separator
        if self.footnote_separator:
            cursor.insertBlock()
            separator_format = QTextBlockFormat()
            separator_format.setBottomMargin(10)
            cursor.setBlockFormat(separator_format)
            cursor.insertText("_" * 50)  # Separator line

        # Insert footnotes
        cursor.insertBlock()
        for note in sorted(footnotes, key=lambda n: n.position):
            self._insert_note_text(cursor, note)

    def insert_endnotes_section(self):
        """Insert endnotes section at the end of the document."""
        endnotes = self.get_endnotes()
        if not endnotes:
            return

        # Move to end of document
        cursor = self.parent.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)

        # Insert section title
        cursor.insertBlock()
        title_format = QTextCharFormat()
        title_format.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        cursor.setCharFormat(title_format)
        cursor.insertText("Endnotes")

        # Insert endnotes
        cursor.insertBlock()
        for note in sorted(endnotes, key=lambda n: n.position):
            self._insert_note_text(cursor, note)

    def _insert_note_text(self, cursor, note):
        """Insert note text at cursor position."""
        # Insert note number/mark
        mark_format = QTextCharFormat()
        mark_format.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSuperScript)
        mark_format.setFontPointSize(8)

        cursor.setCharFormat(mark_format)
        mark_text = note.reference_mark if note.reference_mark else note.number
        cursor.insertText(str(mark_text))

        # Insert note text
        text_format = QTextCharFormat()
        text_format.setFontPointSize(10)
        cursor.setCharFormat(text_format)
        cursor.insertText(f" {note.text}")
        cursor.insertBlock()

    def navigate_to_reference(self, note_id):
        """Navigate to the reference point of a note."""
        note = self.get_note_by_id(note_id)
        if note:
            cursor = self.parent.textCursor()
            cursor.setPosition(note.position)
            self.parent.setTextCursor(cursor)
            self.parent.ensureCursorVisible()
            return True
        return False

    def convert_footnote_to_endnote(self, note_id):
        """Convert a footnote to an endnote."""
        note = self.get_note_by_id(note_id)
        if note and note.note_type == 'footnote':
            note.note_type = 'endnote'
            self._renumber_notes()
            return True
        return False

    def convert_endnote_to_footnote(self, note_id):
        """Convert an endnote to a footnote."""
        note = self.get_note_by_id(note_id)
        if note and note.note_type == 'endnote':
            note.note_type = 'footnote'
            self._renumber_notes()
            return True
        return False


class FootnoteDialog(QDialog):
    """Dialog for adding/editing footnotes and endnotes."""

    def __init__(self, note=None, note_type='footnote', parent=None):
        super().__init__(parent)
        self.note = note
        self.note_type = note_type

        title = "Edit" if note else "Insert"
        title += " Footnote" if note_type == 'footnote' else " Endnote"
        self.setWindowTitle(title)

        self.setModal(True)
        self.setMinimumWidth(500)
        self.setMinimumHeight(300)

        self.setup_ui()

        if note:
            self.text_edit.setPlainText(note.text)
            if note.reference_mark:
                self.custom_mark_edit.setText(note.reference_mark)

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Text input
        text_label = QLabel("Note text:")
        layout.addWidget(text_label)

        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Enter note text here...")
        self.text_edit.setMinimumHeight(150)
        layout.addWidget(self.text_edit)

        # Custom reference mark
        mark_layout = QHBoxLayout()
        mark_layout.addWidget(QLabel("Custom reference mark (optional):"))
        self.custom_mark_edit = QTextEdit()
        self.custom_mark_edit.setMaximumHeight(30)
        self.custom_mark_edit.setPlaceholderText("e.g., *, †, or leave empty for auto-numbering")
        mark_layout.addWidget(self.custom_mark_edit)
        layout.addLayout(mark_layout)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def get_text(self):
        """Get the note text."""
        return self.text_edit.toPlainText()

    def get_custom_mark(self):
        """Get the custom reference mark."""
        mark = self.custom_mark_edit.toPlainText().strip()
        return mark if mark else None


class FootnotesViewer(QDialog):
    """Dialog for viewing and managing footnotes and endnotes."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Footnotes and Endnotes")
        self.setModal(False)
        self.setMinimumSize(700, 500)

        self.setup_ui()
        self.refresh_notes()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Header
        header_label = QLabel("<h3>Footnotes and Endnotes</h3>")
        layout.addWidget(header_label)

        # Filter
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Show:"))

        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["All Notes", "Footnotes", "Endnotes"])
        self.filter_combo.currentTextChanged.connect(self.refresh_notes)
        filter_layout.addWidget(self.filter_combo)
        filter_layout.addStretch()

        layout.addLayout(filter_layout)

        # Notes list
        self.notes_list = QListWidget()
        self.notes_list.itemSelectionChanged.connect(self.on_note_selected)
        self.notes_list.itemDoubleClicked.connect(self.navigate_to_note)
        layout.addWidget(self.notes_list)

        # Note preview
        preview_label = QLabel("<b>Preview:</b>")
        layout.addWidget(preview_label)

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(100)
        layout.addWidget(self.preview_text)

        # Action buttons
        button_layout = QHBoxLayout()

        self.edit_button = QPushButton("Edit")
        self.edit_button.clicked.connect(self.edit_note)
        button_layout.addWidget(self.edit_button)

        self.delete_button = QPushButton("Delete")
        self.delete_button.clicked.connect(self.delete_note)
        button_layout.addWidget(self.delete_button)

        self.convert_button = QPushButton("Convert")
        self.convert_button.clicked.connect(self.convert_note)
        button_layout.addWidget(self.convert_button)

        self.goto_button = QPushButton("Go To Reference")
        self.goto_button.clicked.connect(self.navigate_to_note)
        button_layout.addWidget(self.goto_button)

        button_layout.addStretch()

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def refresh_notes(self):
        """Refresh the notes list."""
        self.notes_list.clear()
        self.preview_text.clear()

        filter_text = self.filter_combo.currentText()

        if filter_text == "All Notes":
            notes = self.manager.notes
        elif filter_text == "Footnotes":
            notes = self.manager.get_footnotes()
        elif filter_text == "Endnotes":
            notes = self.manager.get_endnotes()
        else:
            notes = self.manager.notes

        # Sort by position
        notes = sorted(notes, key=lambda n: n.position)

        for note in notes:
            mark = note.reference_mark if note.reference_mark else note.number
            note_type = "Footnote" if note.note_type == 'footnote' else "Endnote"
            preview = note.text[:50] + "..." if len(note.text) > 50 else note.text

            item_text = f"[{mark}] {note_type}: {preview}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, note.id)
            self.notes_list.addItem(item)

    def on_note_selected(self):
        """Handle note selection."""
        items = self.notes_list.selectedItems()
        if not items:
            self.preview_text.clear()
            return

        note_id = items[0].data(Qt.ItemDataRole.UserRole)
        note = self.manager.get_note_by_id(note_id)

        if note:
            self.preview_text.setPlainText(note.text)

    def edit_note(self):
        """Edit the selected note."""
        items = self.notes_list.selectedItems()
        if not items:
            return

        note_id = items[0].data(Qt.ItemDataRole.UserRole)
        note = self.manager.get_note_by_id(note_id)

        if note:
            dialog = FootnoteDialog(note, note.note_type, parent=self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                new_text = dialog.get_text()
                custom_mark = dialog.get_custom_mark()

                self.manager.edit_note(note_id, new_text)
                note.reference_mark = custom_mark

                self.refresh_notes()

    def delete_note(self):
        """Delete the selected note."""
        items = self.notes_list.selectedItems()
        if not items:
            return

        reply = QMessageBox.question(
            self,
            "Delete Note",
            "Are you sure you want to delete this note?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            note_id = items[0].data(Qt.ItemDataRole.UserRole)
            self.manager.delete_note(note_id)
            self.refresh_notes()

    def convert_note(self):
        """Convert between footnote and endnote."""
        items = self.notes_list.selectedItems()
        if not items:
            return

        note_id = items[0].data(Qt.ItemDataRole.UserRole)
        note = self.manager.get_note_by_id(note_id)

        if note:
            if note.note_type == 'footnote':
                self.manager.convert_footnote_to_endnote(note_id)
                QMessageBox.information(self, "Converted", "Footnote converted to endnote.")
            else:
                self.manager.convert_endnote_to_footnote(note_id)
                QMessageBox.information(self, "Converted", "Endnote converted to footnote.")

            self.refresh_notes()

    def navigate_to_note(self):
        """Navigate to the note reference in the document."""
        items = self.notes_list.selectedItems()
        if not items:
            return

        note_id = items[0].data(Qt.ItemDataRole.UserRole)
        self.manager.navigate_to_reference(note_id)


class FootnoteOptionsDialog(QDialog):
    """Dialog for configuring footnote and endnote options."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Footnote and Endnote Options")
        self.setModal(True)
        self.setMinimumWidth(500)

        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Footnote settings
        footnote_group = QGroupBox("Footnote Settings")
        footnote_layout = QFormLayout()

        self.footnote_numbering_combo = QComboBox()
        self.footnote_numbering_combo.addItems(["Numeric (1, 2, 3...)", "Alphabetic (a, b, c...)",
                                                "Roman (i, ii, iii...)", "Symbols (*, †, ‡...)"])
        footnote_layout.addRow("Numbering style:", self.footnote_numbering_combo)

        self.footnote_start_spin = QSpinBox()
        self.footnote_start_spin.setMinimum(1)
        self.footnote_start_spin.setMaximum(1000)
        footnote_layout.addRow("Start at:", self.footnote_start_spin)

        footnote_group.setLayout(footnote_layout)
        layout.addWidget(footnote_group)

        # Endnote settings
        endnote_group = QGroupBox("Endnote Settings")
        endnote_layout = QFormLayout()

        self.endnote_numbering_combo = QComboBox()
        self.endnote_numbering_combo.addItems(["Numeric (1, 2, 3...)", "Alphabetic (a, b, c...)",
                                               "Roman (i, ii, iii...)", "Symbols (*, †, ‡...)"])
        endnote_layout.addRow("Numbering style:", self.endnote_numbering_combo)

        self.endnote_start_spin = QSpinBox()
        self.endnote_start_spin.setMinimum(1)
        self.endnote_start_spin.setMaximum(1000)
        endnote_layout.addRow("Start at:", self.endnote_start_spin)

        endnote_group.setLayout(endnote_layout)
        layout.addWidget(endnote_group)

        # General settings
        general_group = QGroupBox("General Settings")
        general_layout = QVBoxLayout()

        self.restart_combo = QComboBox()
        self.restart_combo.addItems(["Continuous", "Each Page", "Each Section"])
        restart_layout = QHBoxLayout()
        restart_layout.addWidget(QLabel("Restart numbering:"))
        restart_layout.addWidget(self.restart_combo)
        restart_layout.addStretch()
        general_layout.addLayout(restart_layout)

        self.separator_check = QCheckBox("Show separator line")
        general_layout.addWidget(self.separator_check)

        general_group.setLayout(general_layout)
        layout.addWidget(general_group)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.save_settings)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def load_settings(self):
        """Load current settings into dialog."""
        # Map numbering style to combo index
        numbering_map = {
            'numeric': 0,
            'alphabetic': 1,
            'roman': 2,
            'symbols': 3
        }

        self.footnote_numbering_combo.setCurrentIndex(
            numbering_map.get(self.manager.footnote_numbering, 0)
        )
        self.endnote_numbering_combo.setCurrentIndex(
            numbering_map.get(self.manager.endnote_numbering, 0)
        )

        self.footnote_start_spin.setValue(self.manager.footnote_start_number)
        self.endnote_start_spin.setValue(self.manager.endnote_start_number)

        restart_map = {
            'continuous': 0,
            'each_page': 1,
            'each_section': 2
        }
        self.restart_combo.setCurrentIndex(
            restart_map.get(self.manager.restart_numbering, 0)
        )

        self.separator_check.setChecked(self.manager.footnote_separator)

    def save_settings(self):
        """Save settings from dialog."""
        # Map combo index to numbering style
        numbering_styles = ['numeric', 'alphabetic', 'roman', 'symbols']

        self.manager.footnote_numbering = numbering_styles[self.footnote_numbering_combo.currentIndex()]
        self.manager.endnote_numbering = numbering_styles[self.endnote_numbering_combo.currentIndex()]

        self.manager.footnote_start_number = self.footnote_start_spin.value()
        self.manager.endnote_start_number = self.endnote_start_spin.value()

        restart_styles = ['continuous', 'each_page', 'each_section']
        self.manager.restart_numbering = restart_styles[self.restart_combo.currentIndex()]

        self.manager.footnote_separator = self.separator_check.isChecked()

        # Renumber all notes
        self.manager._renumber_notes()

        self.accept()
