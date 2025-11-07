"""
Captions System for PyWord.

This module provides comprehensive support for captions on figures, tables,
equations, and other objects, with automatic numbering and formatting.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QListWidget, QListWidgetItem, QGroupBox,
                               QComboBox, QTextEdit, QFormLayout, QMessageBox,
                               QCheckBox, QSpinBox, QRadioButton, QButtonGroup)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QTextCursor, QTextCharFormat, QFont, QColor, QTextBlockFormat
from datetime import datetime
import uuid


class Caption:
    """Represents a caption for a figure, table, or other object."""

    def __init__(self, text, caption_type='figure', position=0):
        self.id = str(uuid.uuid4())
        self.text = text
        self.caption_type = caption_type  # 'figure', 'table', 'equation', 'listing'
        self.position = position  # Position in document
        self.number = None  # Auto-assigned number
        self.label = None  # Optional custom label for cross-referencing
        self.chapter_number = None  # For chapter-based numbering
        self.created = datetime.now()

    def to_dict(self):
        """Convert caption to dictionary for serialization."""
        return {
            'id': self.id,
            'text': self.text,
            'caption_type': self.caption_type,
            'position': self.position,
            'number': self.number,
            'label': self.label,
            'chapter_number': self.chapter_number,
            'created': self.created.isoformat()
        }

    @staticmethod
    def from_dict(data):
        """Create Caption from dictionary."""
        caption = Caption(
            data['text'],
            data.get('caption_type', 'figure'),
            data.get('position', 0)
        )
        caption.id = data['id']
        caption.number = data.get('number')
        caption.label = data.get('label')
        caption.chapter_number = data.get('chapter_number')
        if 'created' in data:
            caption.created = datetime.fromisoformat(data['created'])
        return caption

    def get_formatted_number(self, numbering_style='arabic', include_chapter=False):
        """Get formatted caption number."""
        if self.number is None:
            return ""

        # Format number based on style
        if numbering_style == 'arabic':
            num_str = str(self.number)
        elif numbering_style == 'roman':
            num_str = self._to_roman(self.number)
        elif numbering_style == 'alphabetic':
            num_str = self._to_alphabetic(self.number)
        else:
            num_str = str(self.number)

        # Add chapter number if needed
        if include_chapter and self.chapter_number:
            return f"{self.chapter_number}.{num_str}"

        return num_str

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

        return result

    def _to_alphabetic(self, number):
        """Convert number to alphabetic (A, B, C, ...)."""
        result = ""
        while number > 0:
            number -= 1
            result = chr(ord('A') + (number % 26)) + result
            number //= 26
        return result


class CaptionsManager:
    """Manages captions in a document."""

    def __init__(self, parent):
        self.parent = parent  # Reference to text editor
        self.captions = []

        # Settings
        self.numbering_style = 'arabic'  # 'arabic', 'roman', 'alphabetic'
        self.include_chapter = False
        self.restart_each_chapter = False
        self.caption_position = 'below'  # 'above' or 'below' object

        # Prefixes for different caption types
        self.prefixes = {
            'figure': 'Figure',
            'table': 'Table',
            'equation': 'Equation',
            'listing': 'Listing'
        }

        # Formatting
        self.caption_font_size = 10
        self.caption_font_italic = True
        self.caption_alignment = 'center'  # 'left', 'center', 'right'
        self.separator = ': '  # Separator between number and text

    def add_caption(self, text, caption_type='figure', position=None, label=None):
        """Add a new caption."""
        if position is None:
            cursor = self.parent.textCursor()
            position = cursor.position()

        caption = Caption(text, caption_type, position)
        caption.label = label

        self.captions.append(caption)
        self._renumber_captions()

        return caption

    def edit_caption(self, caption_id, new_text):
        """Edit an existing caption."""
        caption = self.get_caption_by_id(caption_id)
        if caption:
            caption.text = new_text
            return True
        return False

    def delete_caption(self, caption_id):
        """Delete a caption."""
        caption = self.get_caption_by_id(caption_id)
        if caption:
            self.captions.remove(caption)
            self._renumber_captions()
            return True
        return False

    def get_caption_by_id(self, caption_id):
        """Get a caption by its ID."""
        for caption in self.captions:
            if caption.id == caption_id:
                return caption
        return None

    def get_caption_by_label(self, label):
        """Get a caption by its label."""
        for caption in self.captions:
            if caption.label == label:
                return caption
        return None

    def get_captions_by_type(self, caption_type):
        """Get all captions of a specific type."""
        return [c for c in self.captions if c.caption_type == caption_type]

    def _renumber_captions(self):
        """Renumber all captions based on their position and type."""
        # Group captions by type
        by_type = {}
        for caption in self.captions:
            if caption.caption_type not in by_type:
                by_type[caption.caption_type] = []
            by_type[caption.caption_type].append(caption)

        # Sort each type by position and renumber
        for caption_type, captions_list in by_type.items():
            sorted_captions = sorted(captions_list, key=lambda c: c.position)

            current_chapter = 1
            chapter_counter = 1

            for i, caption in enumerate(sorted_captions):
                # Simplified chapter detection
                # In a real implementation, you'd detect actual chapter breaks
                caption.chapter_number = current_chapter

                if self.restart_each_chapter:
                    caption.number = chapter_counter
                else:
                    caption.number = i + 1

    def insert_caption(self, text, caption_type='figure', position=None, label=None):
        """Insert a formatted caption at the specified position."""
        if position is None:
            cursor = self.parent.textCursor()
        else:
            cursor = QTextCursor(self.parent.document())
            cursor.setPosition(position)

        # Add caption to manager
        caption = self.add_caption(text, caption_type, cursor.position(), label)

        # Insert formatted caption
        self._insert_formatted_caption(cursor, caption)

        return caption

    def _insert_formatted_caption(self, cursor, caption):
        """Insert a formatted caption at cursor position."""
        # Move to new line if needed
        cursor.insertBlock()

        # Set block format (alignment)
        block_format = QTextBlockFormat()
        if self.caption_alignment == 'center':
            block_format.setAlignment(Qt.AlignmentFlag.AlignCenter)
        elif self.caption_alignment == 'right':
            block_format.setAlignment(Qt.AlignmentFlag.AlignRight)
        else:
            block_format.setAlignment(Qt.AlignmentFlag.AlignLeft)

        cursor.setBlockFormat(block_format)

        # Set character format
        char_format = QTextCharFormat()
        char_format.setFontPointSize(self.caption_font_size)
        char_format.setFontItalic(self.caption_font_italic)

        # Insert prefix and number
        prefix = self.prefixes.get(caption.caption_type, caption.caption_type.capitalize())
        number = caption.get_formatted_number(self.numbering_style, self.include_chapter)

        caption_text = f"{prefix} {number}{self.separator}{caption.text}"

        cursor.setCharFormat(char_format)
        cursor.insertText(caption_text)

    def update_caption(self, caption_id):
        """Update an existing caption in the document."""
        caption = self.get_caption_by_id(caption_id)
        if not caption:
            return False

        # Find and update the caption in the document
        # This is simplified - in a real implementation,
        # you'd track exact positions and update them
        self._renumber_captions()

        return True

    def insert_list_of_figures(self, position=None):
        """Insert a list of all figures."""
        return self._insert_list_of_captions('figure', "List of Figures", position)

    def insert_list_of_tables(self, position=None):
        """Insert a list of all tables."""
        return self._insert_list_of_captions('table', "List of Tables", position)

    def _insert_list_of_captions(self, caption_type, title, position=None):
        """Insert a list of captions of a specific type."""
        if position is None:
            cursor = self.parent.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.End)
        else:
            cursor = QTextCursor(self.parent.document())
            cursor.setPosition(position)

        # Get captions of specified type
        captions = self.get_captions_by_type(caption_type)

        if not captions:
            QMessageBox.warning(
                None,
                f"No {caption_type.capitalize()}s",
                f"No {caption_type}s found in the document."
            )
            return False

        # Insert title
        cursor.insertBlock()
        title_format = QTextCharFormat()
        title_format.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        cursor.setCharFormat(title_format)
        cursor.insertText(title)

        # Insert captions list
        cursor.insertBlock()

        sorted_captions = sorted(captions, key=lambda c: c.position)

        for caption in sorted_captions:
            cursor.insertBlock()

            # Format entry
            entry_format = QTextCharFormat()
            entry_format.setForeground(QColor(0, 0, 255))  # Blue for links
            entry_format.setFontUnderline(True)

            prefix = self.prefixes.get(caption_type, caption_type.capitalize())
            number = caption.get_formatted_number(self.numbering_style, self.include_chapter)

            entry_text = f"{prefix} {number}: {caption.text}"

            cursor.setCharFormat(entry_format)
            cursor.insertText(entry_text)

        return True

    def navigate_to_caption(self, caption_id):
        """Navigate to a caption in the document."""
        caption = self.get_caption_by_id(caption_id)
        if caption:
            cursor = self.parent.textCursor()
            cursor.setPosition(caption.position)
            self.parent.setTextCursor(cursor)
            self.parent.ensureCursorVisible()
            return True
        return False

    def get_caption_reference(self, caption_id):
        """Get a formatted reference to a caption (for cross-references)."""
        caption = self.get_caption_by_id(caption_id)
        if caption:
            prefix = self.prefixes.get(caption.caption_type, caption.caption_type.capitalize())
            number = caption.get_formatted_number(self.numbering_style, self.include_chapter)
            return f"{prefix} {number}"
        return None


class CaptionDialog(QDialog):
    """Dialog for adding/editing a caption."""

    def __init__(self, caption=None, caption_type='figure', parent=None):
        super().__init__(parent)
        self.caption = caption
        self.caption_type = caption_type

        title = "Edit Caption" if caption else "Insert Caption"
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(500)

        self.setup_ui()

        if caption:
            self.text_edit.setPlainText(caption.text)
            self.type_combo.setCurrentText(caption.caption_type)
            if caption.label:
                self.label_edit.setText(caption.label)

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Caption type
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Caption for:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(['figure', 'table', 'equation', 'listing'])
        self.type_combo.setCurrentText(self.caption_type)
        type_layout.addWidget(self.type_combo)
        type_layout.addStretch()
        layout.addLayout(type_layout)

        # Caption text
        text_label = QLabel("Caption text:")
        layout.addWidget(text_label)

        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Enter caption text here...")
        self.text_edit.setMinimumHeight(100)
        layout.addWidget(self.text_edit)

        # Optional label for cross-referencing
        label_layout = QHBoxLayout()
        label_layout.addWidget(QLabel("Label (for cross-referencing):"))
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("e.g., fig:my_figure")
        label_layout.addWidget(self.label_edit)
        layout.addLayout(label_layout)

        # Position options
        position_group = QGroupBox("Position")
        position_layout = QVBoxLayout()

        self.position_group = QButtonGroup()
        self.above_radio = QRadioButton("Above object")
        self.below_radio = QRadioButton("Below object")
        self.below_radio.setChecked(True)

        self.position_group.addButton(self.above_radio)
        self.position_group.addButton(self.below_radio)

        position_layout.addWidget(self.above_radio)
        position_layout.addWidget(self.below_radio)

        position_group.setLayout(position_layout)
        layout.addWidget(position_group)

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

    def get_caption_text(self):
        """Get the caption text."""
        return self.text_edit.toPlainText()

    def get_caption_type(self):
        """Get the caption type."""
        return self.type_combo.currentText()

    def get_label(self):
        """Get the optional label."""
        label = self.label_edit.text().strip()
        return label if label else None

    def get_position(self):
        """Get the caption position."""
        return 'above' if self.above_radio.isChecked() else 'below'


class CaptionsViewer(QDialog):
    """Dialog for viewing and managing captions."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Captions")
        self.setModal(False)
        self.setMinimumSize(700, 500)

        self.setup_ui()
        self.refresh_captions()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Header
        header_label = QLabel("<h3>Captions</h3>")
        layout.addWidget(header_label)

        # Filter
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Show:"))

        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["All Captions", "Figures", "Tables", "Equations", "Listings"])
        self.filter_combo.currentTextChanged.connect(self.refresh_captions)
        filter_layout.addWidget(self.filter_combo)
        filter_layout.addStretch()

        layout.addLayout(filter_layout)

        # Captions list
        self.captions_list = QListWidget()
        self.captions_list.itemSelectionChanged.connect(self.on_caption_selected)
        self.captions_list.itemDoubleClicked.connect(self.navigate_to_caption)
        layout.addWidget(self.captions_list)

        # Caption preview
        preview_label = QLabel("<b>Preview:</b>")
        layout.addWidget(preview_label)

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(80)
        layout.addWidget(self.preview_text)

        # Action buttons
        button_layout = QHBoxLayout()

        self.edit_button = QPushButton("Edit")
        self.edit_button.clicked.connect(self.edit_caption)
        button_layout.addWidget(self.edit_button)

        self.delete_button = QPushButton("Delete")
        self.delete_button.clicked.connect(self.delete_caption)
        button_layout.addWidget(self.delete_button)

        self.goto_button = QPushButton("Go To")
        self.goto_button.clicked.connect(self.navigate_to_caption)
        button_layout.addWidget(self.goto_button)

        button_layout.addStretch()

        # List generation buttons
        list_figures_button = QPushButton("List of Figures")
        list_figures_button.clicked.connect(self.insert_list_of_figures)
        button_layout.addWidget(list_figures_button)

        list_tables_button = QPushButton("List of Tables")
        list_tables_button.clicked.connect(self.insert_list_of_tables)
        button_layout.addWidget(list_tables_button)

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def refresh_captions(self):
        """Refresh the captions list."""
        self.captions_list.clear()
        self.preview_text.clear()

        filter_text = self.filter_combo.currentText()

        if filter_text == "All Captions":
            captions = self.manager.captions
        elif filter_text == "Figures":
            captions = self.manager.get_captions_by_type('figure')
        elif filter_text == "Tables":
            captions = self.manager.get_captions_by_type('table')
        elif filter_text == "Equations":
            captions = self.manager.get_captions_by_type('equation')
        elif filter_text == "Listings":
            captions = self.manager.get_captions_by_type('listing')
        else:
            captions = self.manager.captions

        # Sort by position
        captions = sorted(captions, key=lambda c: c.position)

        for caption in captions:
            prefix = self.manager.prefixes.get(caption.caption_type, caption.caption_type.capitalize())
            number = caption.get_formatted_number(
                self.manager.numbering_style,
                self.manager.include_chapter
            )

            preview = caption.text[:40] + "..." if len(caption.text) > 40 else caption.text

            item_text = f"{prefix} {number}: {preview}"
            if caption.label:
                item_text += f" [{caption.label}]"

            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, caption.id)
            self.captions_list.addItem(item)

    def on_caption_selected(self):
        """Handle caption selection."""
        items = self.captions_list.selectedItems()
        if not items:
            self.preview_text.clear()
            return

        caption_id = items[0].data(Qt.ItemDataRole.UserRole)
        caption = self.manager.get_caption_by_id(caption_id)

        if caption:
            self.preview_text.setPlainText(caption.text)

    def edit_caption(self):
        """Edit the selected caption."""
        items = self.captions_list.selectedItems()
        if not items:
            return

        caption_id = items[0].data(Qt.ItemDataRole.UserRole)
        caption = self.manager.get_caption_by_id(caption_id)

        if caption:
            dialog = CaptionDialog(caption, caption.caption_type, parent=self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                new_text = dialog.get_caption_text()
                new_label = dialog.get_label()

                self.manager.edit_caption(caption_id, new_text)
                caption.label = new_label
                caption.caption_type = dialog.get_caption_type()

                self.manager._renumber_captions()
                self.refresh_captions()

    def delete_caption(self):
        """Delete the selected caption."""
        items = self.captions_list.selectedItems()
        if not items:
            return

        reply = QMessageBox.question(
            self,
            "Delete Caption",
            "Are you sure you want to delete this caption?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            caption_id = items[0].data(Qt.ItemDataRole.UserRole)
            self.manager.delete_caption(caption_id)
            self.refresh_captions()

    def navigate_to_caption(self):
        """Navigate to the caption in the document."""
        items = self.captions_list.selectedItems()
        if not items:
            return

        caption_id = items[0].data(Qt.ItemDataRole.UserRole)
        self.manager.navigate_to_caption(caption_id)

    def insert_list_of_figures(self):
        """Insert list of figures."""
        if self.manager.insert_list_of_figures():
            QMessageBox.information(self, "Success", "List of figures inserted successfully!")

    def insert_list_of_tables(self):
        """Insert list of tables."""
        if self.manager.insert_list_of_tables():
            QMessageBox.information(self, "Success", "List of tables inserted successfully!")


class CaptionOptionsDialog(QDialog):
    """Dialog for configuring caption options."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Caption Options")
        self.setModal(True)
        self.setMinimumWidth(500)

        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Numbering group
        numbering_group = QGroupBox("Numbering")
        numbering_layout = QFormLayout()

        self.numbering_combo = QComboBox()
        self.numbering_combo.addItems(["Arabic (1, 2, 3...)", "Roman (I, II, III...)", "Alphabetic (A, B, C...)"])
        numbering_layout.addRow("Number style:", self.numbering_combo)

        self.include_chapter_cb = QCheckBox("Include chapter number")
        numbering_layout.addRow("", self.include_chapter_cb)

        self.restart_chapter_cb = QCheckBox("Restart numbering each chapter")
        numbering_layout.addRow("", self.restart_chapter_cb)

        numbering_group.setLayout(numbering_layout)
        layout.addWidget(numbering_group)

        # Formatting group
        formatting_group = QGroupBox("Formatting")
        formatting_layout = QFormLayout()

        self.font_size_spin = QSpinBox()
        self.font_size_spin.setMinimum(6)
        self.font_size_spin.setMaximum(72)
        formatting_layout.addRow("Font size:", self.font_size_spin)

        self.italic_cb = QCheckBox("Italic")
        formatting_layout.addRow("", self.italic_cb)

        self.alignment_combo = QComboBox()
        self.alignment_combo.addItems(["Left", "Center", "Right"])
        formatting_layout.addRow("Alignment:", self.alignment_combo)

        self.separator_edit = QLineEdit()
        self.separator_edit.setMaxLength(5)
        formatting_layout.addRow("Separator:", self.separator_edit)

        formatting_group.setLayout(formatting_layout)
        layout.addWidget(formatting_group)

        # Position group
        position_group = QGroupBox("Position")
        position_layout = QVBoxLayout()

        self.position_combo = QComboBox()
        self.position_combo.addItems(["Below object", "Above object"])
        position_layout.addWidget(self.position_combo)

        position_group.setLayout(position_layout)
        layout.addWidget(position_group)

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
        # Numbering style
        numbering_map = {
            'arabic': 0,
            'roman': 1,
            'alphabetic': 2
        }
        self.numbering_combo.setCurrentIndex(
            numbering_map.get(self.manager.numbering_style, 0)
        )

        self.include_chapter_cb.setChecked(self.manager.include_chapter)
        self.restart_chapter_cb.setChecked(self.manager.restart_each_chapter)

        # Formatting
        self.font_size_spin.setValue(self.manager.caption_font_size)
        self.italic_cb.setChecked(self.manager.caption_font_italic)

        alignment_map = {
            'left': 0,
            'center': 1,
            'right': 2
        }
        self.alignment_combo.setCurrentIndex(
            alignment_map.get(self.manager.caption_alignment, 1)
        )

        self.separator_edit.setText(self.manager.separator)

        # Position
        position_map = {
            'below': 0,
            'above': 1
        }
        self.position_combo.setCurrentIndex(
            position_map.get(self.manager.caption_position, 0)
        )

    def save_settings(self):
        """Save settings from dialog."""
        # Numbering style
        numbering_styles = ['arabic', 'roman', 'alphabetic']
        self.manager.numbering_style = numbering_styles[self.numbering_combo.currentIndex()]

        self.manager.include_chapter = self.include_chapter_cb.isChecked()
        self.manager.restart_each_chapter = self.restart_chapter_cb.isChecked()

        # Formatting
        self.manager.caption_font_size = self.font_size_spin.value()
        self.manager.caption_font_italic = self.italic_cb.isChecked()

        alignment_styles = ['left', 'center', 'right']
        self.manager.caption_alignment = alignment_styles[self.alignment_combo.currentIndex()]

        self.manager.separator = self.separator_edit.text()

        # Position
        position_styles = ['below', 'above']
        self.manager.caption_position = position_styles[self.position_combo.currentIndex()]

        # Renumber all captions
        self.manager._renumber_captions()

        self.accept()
