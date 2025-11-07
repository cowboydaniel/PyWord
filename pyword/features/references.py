"""
Table of Contents System for PyWord.

This module provides a comprehensive table of contents system for documents,
including automatic heading detection, TOC generation, and navigation.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                               QListWidget, QListWidgetItem, QGroupBox, QCheckBox,
                               QSpinBox, QComboBox, QMessageBox, QTreeWidget, QTreeWidgetItem)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QTextCursor, QTextCharFormat, QTextBlockFormat, QFont, QColor
import re


class TOCEntry:
    """Represents a single entry in the table of contents."""

    def __init__(self, text, level, position, page_number=1):
        self.text = text
        self.level = level  # Heading level (1-9)
        self.position = position  # Position in document
        self.page_number = page_number
        self.children = []  # Child entries

    def to_dict(self):
        """Convert to dictionary for serialization."""
        return {
            'text': self.text,
            'level': self.level,
            'position': self.position,
            'page_number': self.page_number,
            'children': [child.to_dict() for child in self.children]
        }

    @staticmethod
    def from_dict(data):
        """Create TOCEntry from dictionary."""
        entry = TOCEntry(
            data['text'],
            data['level'],
            data['position'],
            data.get('page_number', 1)
        )
        entry.children = [TOCEntry.from_dict(child) for child in data.get('children', [])]
        return entry


class TableOfContentsManager:
    """Manages table of contents for a document."""

    def __init__(self, parent):
        self.parent = parent  # Reference to text editor
        self.entries = []
        self.auto_update = True
        self.show_page_numbers = True
        self.max_level = 3  # Maximum heading level to include

    def scan_document(self):
        """Scan the document for headings and build TOC entries."""
        self.entries = []
        document = self.parent.document()
        cursor = QTextCursor(document)

        cursor.movePosition(QTextCursor.MoveOperation.Start)
        block = document.firstBlock()
        position = 0

        while block.isValid():
            block_format = block.blockFormat()
            char_format = block.charFormat()

            # Check if block is a heading
            heading_level = self._get_heading_level(block)

            if heading_level > 0 and heading_level <= self.max_level:
                text = block.text().strip()
                if text:
                    # Calculate approximate page number (simplified)
                    page_number = position // 3000 + 1  # Rough estimate

                    entry = TOCEntry(text, heading_level, block.position(), page_number)
                    self.entries.append(entry)

            position += len(block.text())
            block = block.next()

        # Build hierarchical structure
        self._build_hierarchy()

        return self.entries

    def _get_heading_level(self, block):
        """Determine if a block is a heading and its level."""
        # Check block format for heading property
        block_format = block.blockFormat()
        char_format = block.charFormat()

        # Check font size - larger fonts are likely headings
        font_size = char_format.font().pointSize()
        is_bold = char_format.font().bold()

        # Check for heading markers (e.g., "# Heading", "## Heading")
        text = block.text().strip()
        if text.startswith('#'):
            level = 0
            for char in text:
                if char == '#':
                    level += 1
                else:
                    break
            return min(level, 9)

        # Check for style-based headings
        # This is simplified - in a real implementation, you'd check actual styles
        if is_bold and font_size >= 18:
            return 1
        elif is_bold and font_size >= 16:
            return 2
        elif is_bold and font_size >= 14:
            return 3
        elif font_size >= 12 and is_bold:
            return 4

        # Check for heading property in block format
        heading_prop = block_format.property(QTextBlockFormat.Property.UserProperty)
        if heading_prop:
            try:
                return int(heading_prop)
            except (ValueError, TypeError):
                pass

        return 0

    def _build_hierarchy(self):
        """Build hierarchical structure of TOC entries."""
        if not self.entries:
            return

        # Create a stack to track parent entries at each level
        stack = [None] * 10  # Support up to 9 heading levels
        root_entries = []

        for entry in self.entries:
            level = entry.level

            # Find parent entry
            parent = None
            for i in range(level - 1, 0, -1):
                if stack[i] is not None:
                    parent = stack[i]
                    break

            # Add to parent or root
            if parent:
                parent.children.append(entry)
            else:
                root_entries.append(entry)

            # Update stack
            stack[level] = entry
            # Clear lower levels
            for i in range(level + 1, 10):
                stack[i] = None

        # Replace flat list with hierarchical structure
        self.entries = root_entries

    def insert_toc(self, cursor_position=None, style='standard'):
        """Insert a table of contents at the specified position."""
        if cursor_position is None:
            cursor = self.parent.textCursor()
        else:
            cursor = QTextCursor(self.parent.document())
            cursor.setPosition(cursor_position)

        # Scan document for headings
        self.scan_document()

        if not self.entries:
            QMessageBox.warning(
                None,
                "No Headings",
                "No headings found in the document. Please add headings to generate a table of contents."
            )
            return False

        # Insert TOC title
        cursor.insertBlock()
        title_format = QTextCharFormat()
        title_format.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        cursor.setCharFormat(title_format)
        cursor.insertText("Table of Contents\n\n")

        # Insert TOC entries
        self._insert_toc_entries(cursor, self.entries, style)

        return True

    def _insert_toc_entries(self, cursor, entries, style, indent_level=0):
        """Recursively insert TOC entries."""
        for entry in entries:
            # Create entry format
            entry_format = QTextCharFormat()
            entry_format.setForeground(QColor(0, 0, 255))  # Blue for links
            entry_format.setFontUnderline(True)
            entry_format.setAnchor(True)
            entry_format.setAnchorHref(f"#pos_{entry.position}")

            # Create block format with indentation
            block_format = QTextBlockFormat()
            block_format.setLeftMargin(indent_level * 20)

            cursor.insertBlock(block_format)
            cursor.setCharFormat(entry_format)

            # Format entry text based on style
            if style == 'standard':
                indent = "  " * indent_level
                if self.show_page_numbers:
                    entry_text = f"{indent}{entry.text} {'.' * (60 - len(entry.text) - len(indent) - len(str(entry.page_number)))} {entry.page_number}"
                else:
                    entry_text = f"{indent}{entry.text}"
            elif style == 'simple':
                entry_text = f"{'  ' * indent_level}{entry.text}"
            elif style == 'numbered':
                entry_text = f"{'  ' * indent_level}{entry.level}. {entry.text}"

            cursor.insertText(entry_text)

            # Insert children
            if entry.children:
                self._insert_toc_entries(cursor, entry.children, style, indent_level + 1)

    def update_toc(self):
        """Update an existing table of contents."""
        # Find TOC in document
        document = self.parent.document()
        cursor = QTextCursor(document)

        # Search for "Table of Contents" text
        cursor = document.find("Table of Contents")

        if cursor.isNull():
            QMessageBox.warning(
                None,
                "TOC Not Found",
                "No table of contents found. Please insert one first."
            )
            return False

        # Find the start and end of TOC
        toc_start = cursor.position() - len("Table of Contents")

        # Find next major heading or end of relevant section
        cursor.movePosition(QTextCursor.MoveOperation.Down, QTextCursor.MoveMode.MoveAnchor, 20)
        toc_end = cursor.position()

        # Delete old TOC
        cursor.setPosition(toc_start)
        cursor.setPosition(toc_end, QTextCursor.MoveMode.KeepAnchor)
        cursor.removeSelectedText()

        # Insert updated TOC
        cursor.setPosition(toc_start)
        return self.insert_toc(toc_start)

    def navigate_to_entry(self, position):
        """Navigate to a TOC entry in the document."""
        cursor = self.parent.textCursor()
        cursor.setPosition(position)
        self.parent.setTextCursor(cursor)
        self.parent.ensureCursorVisible()

    def get_entries_flat(self):
        """Get a flat list of all TOC entries."""
        flat_list = []

        def flatten(entries):
            for entry in entries:
                flat_list.append(entry)
                if entry.children:
                    flatten(entry.children)

        flatten(self.entries)
        return flat_list

    def export_toc(self, format='text'):
        """Export table of contents in various formats."""
        self.scan_document()

        if format == 'text':
            return self._export_text()
        elif format == 'html':
            return self._export_html()
        elif format == 'markdown':
            return self._export_markdown()

        return ""

    def _export_text(self):
        """Export TOC as plain text."""
        lines = ["Table of Contents\n"]

        def process_entries(entries, level=0):
            for entry in entries:
                indent = "  " * level
                if self.show_page_numbers:
                    lines.append(f"{indent}{entry.text} ... {entry.page_number}")
                else:
                    lines.append(f"{indent}{entry.text}")
                if entry.children:
                    process_entries(entry.children, level + 1)

        process_entries(self.entries)
        return "\n".join(lines)

    def _export_html(self):
        """Export TOC as HTML."""
        html = ["<div class='table-of-contents'>", "<h2>Table of Contents</h2>", "<ul>"]

        def process_entries(entries):
            for entry in entries:
                html.append(f"<li><a href='#pos_{entry.position}'>{entry.text}</a>")
                if entry.children:
                    html.append("<ul>")
                    process_entries(entry.children)
                    html.append("</ul>")
                html.append("</li>")

        process_entries(self.entries)
        html.append("</ul>")
        html.append("</div>")
        return "\n".join(html)

    def _export_markdown(self):
        """Export TOC as Markdown."""
        lines = ["# Table of Contents\n"]

        def process_entries(entries, level=0):
            for entry in entries:
                indent = "  " * level
                lines.append(f"{indent}- [{entry.text}](#pos_{entry.position})")
                if entry.children:
                    process_entries(entry.children, level + 1)

        process_entries(self.entries)
        return "\n".join(lines)


class TOCDialog(QDialog):
    """Dialog for table of contents management."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Table of Contents")
        self.setModal(True)
        self.setMinimumSize(600, 500)

        self.setup_ui()
        self.refresh_preview()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Title
        title_label = QLabel("<h2>Table of Contents</h2>")
        layout.addWidget(title_label)

        # Options group
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()

        # Show page numbers
        self.show_page_numbers_cb = QCheckBox("Show page numbers")
        self.show_page_numbers_cb.setChecked(self.manager.show_page_numbers)
        self.show_page_numbers_cb.toggled.connect(self.on_options_changed)
        options_layout.addWidget(self.show_page_numbers_cb)

        # Auto update
        self.auto_update_cb = QCheckBox("Auto-update on document changes")
        self.auto_update_cb.setChecked(self.manager.auto_update)
        options_layout.addWidget(self.auto_update_cb)

        # Max level
        level_layout = QHBoxLayout()
        level_layout.addWidget(QLabel("Maximum heading level:"))
        self.max_level_spin = QSpinBox()
        self.max_level_spin.setMinimum(1)
        self.max_level_spin.setMaximum(9)
        self.max_level_spin.setValue(self.manager.max_level)
        self.max_level_spin.valueChanged.connect(self.on_options_changed)
        level_layout.addWidget(self.max_level_spin)
        level_layout.addStretch()
        options_layout.addLayout(level_layout)

        # Style
        style_layout = QHBoxLayout()
        style_layout.addWidget(QLabel("Style:"))
        self.style_combo = QComboBox()
        self.style_combo.addItems(["Standard", "Simple", "Numbered"])
        self.style_combo.currentTextChanged.connect(self.on_options_changed)
        style_layout.addWidget(self.style_combo)
        style_layout.addStretch()
        options_layout.addLayout(style_layout)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Preview group
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()

        self.preview_tree = QTreeWidget()
        self.preview_tree.setHeaderLabels(["Heading", "Page"])
        self.preview_tree.itemDoubleClicked.connect(self.on_entry_double_clicked)
        preview_layout.addWidget(self.preview_tree)

        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)

        # Buttons
        button_layout = QHBoxLayout()

        refresh_button = QPushButton("Refresh")
        refresh_button.clicked.connect(self.refresh_preview)
        button_layout.addWidget(refresh_button)

        button_layout.addStretch()

        insert_button = QPushButton("Insert TOC")
        insert_button.clicked.connect(self.insert_toc)
        button_layout.addWidget(insert_button)

        update_button = QPushButton("Update TOC")
        update_button.clicked.connect(self.update_toc)
        button_layout.addWidget(update_button)

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def on_options_changed(self):
        """Handle options changes."""
        self.manager.show_page_numbers = self.show_page_numbers_cb.isChecked()
        self.manager.auto_update = self.auto_update_cb.isChecked()
        self.manager.max_level = self.max_level_spin.value()
        self.refresh_preview()

    def refresh_preview(self):
        """Refresh the TOC preview."""
        self.preview_tree.clear()
        self.manager.scan_document()

        self._add_entries_to_tree(self.manager.entries, self.preview_tree)
        self.preview_tree.expandAll()

    def _add_entries_to_tree(self, entries, parent):
        """Recursively add entries to tree widget."""
        for entry in entries:
            item = QTreeWidgetItem(parent)
            item.setText(0, entry.text)
            item.setText(1, str(entry.page_number) if self.manager.show_page_numbers else "")
            item.setData(0, Qt.ItemDataRole.UserRole, entry.position)

            if entry.children:
                self._add_entries_to_tree(entry.children, item)

    def on_entry_double_clicked(self, item):
        """Navigate to entry on double-click."""
        position = item.data(0, Qt.ItemDataRole.UserRole)
        if position is not None:
            self.manager.navigate_to_entry(position)

    def insert_toc(self):
        """Insert table of contents."""
        style = self.style_combo.currentText().lower()
        if self.manager.insert_toc(style=style):
            QMessageBox.information(self, "Success", "Table of contents inserted successfully!")
            self.accept()

    def update_toc(self):
        """Update existing table of contents."""
        if self.manager.update_toc():
            QMessageBox.information(self, "Success", "Table of contents updated successfully!")
            self.refresh_preview()
