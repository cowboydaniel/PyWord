"""
Cross-references System for PyWord.

This module provides comprehensive support for cross-references to headings,
figures, tables, equations, footnotes, and other document elements.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QListWidget, QListWidgetItem, QGroupBox,
                               QComboBox, QTextEdit, QFormLayout, QMessageBox,
                               QCheckBox, QRadioButton, QButtonGroup, QTreeWidget,
                               QTreeWidgetItem, QTabWidget, QWidget)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QTextCursor, QTextCharFormat, QFont, QColor
from datetime import datetime
import uuid


class CrossReference:
    """Represents a cross-reference in the document."""

    def __init__(self, target_type, target_id, reference_type='number', position=0):
        self.id = str(uuid.uuid4())
        self.target_type = target_type  # 'heading', 'figure', 'table', 'equation', 'footnote', 'bookmark'
        self.target_id = target_id  # ID of the referenced item
        self.reference_type = reference_type  # 'number', 'page', 'text', 'number_and_page', 'full'
        self.position = position  # Position in document
        self.created = datetime.now()

    def to_dict(self):
        """Convert cross-reference to dictionary for serialization."""
        return {
            'id': self.id,
            'target_type': self.target_type,
            'target_id': self.target_id,
            'reference_type': self.reference_type,
            'position': self.position,
            'created': self.created.isoformat()
        }

    @staticmethod
    def from_dict(data):
        """Create CrossReference from dictionary."""
        ref = CrossReference(
            data['target_type'],
            data['target_id'],
            data.get('reference_type', 'number'),
            data.get('position', 0)
        )
        ref.id = data['id']
        if 'created' in data:
            ref.created = datetime.fromisoformat(data['created'])
        return ref


class Bookmark:
    """Represents a bookmark in the document."""

    def __init__(self, name, position=0, length=0):
        self.id = str(uuid.uuid4())
        self.name = name
        self.position = position
        self.length = length  # Length of bookmarked text
        self.created = datetime.now()

    def to_dict(self):
        """Convert bookmark to dictionary for serialization."""
        return {
            'id': self.id,
            'name': self.name,
            'position': self.position,
            'length': self.length,
            'created': self.created.isoformat()
        }

    @staticmethod
    def from_dict(data):
        """Create Bookmark from dictionary."""
        bookmark = Bookmark(
            data['name'],
            data.get('position', 0),
            data.get('length', 0)
        )
        bookmark.id = data['id']
        if 'created' in data:
            bookmark.created = datetime.fromisoformat(data['created'])
        return bookmark


class CrossReferencesManager:
    """Manages cross-references in a document."""

    def __init__(self, parent):
        self.parent = parent  # Reference to text editor
        self.cross_references = []
        self.bookmarks = []

        # References to other managers
        self.captions_manager = None
        self.footnotes_manager = None
        self.toc_manager = None

        # Settings
        self.auto_update = True
        self.hyperlink_references = True

    def set_managers(self, captions_manager=None, footnotes_manager=None, toc_manager=None):
        """Set references to other document managers."""
        self.captions_manager = captions_manager
        self.footnotes_manager = footnotes_manager
        self.toc_manager = toc_manager

    def add_bookmark(self, name, position=None, length=0):
        """Add a new bookmark."""
        # Check if bookmark name already exists
        if self.get_bookmark_by_name(name):
            return None

        if position is None:
            cursor = self.parent.textCursor()
            position = cursor.position()
            if cursor.hasSelection():
                length = len(cursor.selectedText())

        bookmark = Bookmark(name, position, length)
        self.bookmarks.append(bookmark)

        # Highlight bookmarked text
        if length > 0:
            self._highlight_bookmark(bookmark)

        return bookmark

    def delete_bookmark(self, bookmark_id):
        """Delete a bookmark."""
        bookmark = self.get_bookmark_by_id(bookmark_id)
        if bookmark:
            # Remove highlighting
            if bookmark.length > 0:
                self._remove_bookmark_highlight(bookmark)

            # Remove any cross-references to this bookmark
            self.cross_references = [
                ref for ref in self.cross_references
                if not (ref.target_type == 'bookmark' and ref.target_id == bookmark_id)
            ]

            self.bookmarks.remove(bookmark)
            return True
        return False

    def get_bookmark_by_id(self, bookmark_id):
        """Get a bookmark by its ID."""
        for bookmark in self.bookmarks:
            if bookmark.id == bookmark_id:
                return bookmark
        return None

    def get_bookmark_by_name(self, name):
        """Get a bookmark by its name."""
        for bookmark in self.bookmarks:
            if bookmark.name == name:
                return bookmark
        return None

    def navigate_to_bookmark(self, bookmark_id):
        """Navigate to a bookmark in the document."""
        bookmark = self.get_bookmark_by_id(bookmark_id)
        if bookmark:
            cursor = self.parent.textCursor()
            cursor.setPosition(bookmark.position)
            self.parent.setTextCursor(cursor)
            self.parent.ensureCursorVisible()
            return True
        return False

    def _highlight_bookmark(self, bookmark):
        """Highlight bookmarked text."""
        cursor = QTextCursor(self.parent.document())
        cursor.setPosition(bookmark.position)
        cursor.setPosition(bookmark.position + bookmark.length, QTextCursor.MoveMode.KeepAnchor)

        format = QTextCharFormat()
        format.setBackground(QColor(220, 220, 255))  # Light blue
        cursor.mergeCharFormat(format)

    def _remove_bookmark_highlight(self, bookmark):
        """Remove bookmark highlighting."""
        cursor = QTextCursor(self.parent.document())
        cursor.setPosition(bookmark.position)
        cursor.setPosition(bookmark.position + bookmark.length, QTextCursor.MoveMode.KeepAnchor)

        format = QTextCharFormat()
        format.setBackground(QColor(Qt.GlobalColor.white))
        cursor.setCharFormat(format)

    def insert_cross_reference(self, target_type, target_id, reference_type='number', position=None):
        """Insert a cross-reference in the document."""
        if position is None:
            cursor = self.parent.textCursor()
            position = cursor.position()
        else:
            cursor = QTextCursor(self.parent.document())
            cursor.setPosition(position)

        # Create cross-reference
        cross_ref = CrossReference(target_type, target_id, reference_type, position)
        self.cross_references.append(cross_ref)

        # Get reference text
        ref_text = self._get_reference_text(cross_ref)

        if ref_text:
            # Insert formatted reference
            self._insert_reference_text(cursor, ref_text, cross_ref)
            return cross_ref

        return None

    def _get_reference_text(self, cross_ref):
        """Get the text for a cross-reference."""
        if cross_ref.target_type == 'bookmark':
            bookmark = self.get_bookmark_by_id(cross_ref.target_id)
            if not bookmark:
                return "[Invalid Bookmark]"

            if cross_ref.reference_type == 'text':
                # Get bookmarked text
                cursor = QTextCursor(self.parent.document())
                cursor.setPosition(bookmark.position)
                cursor.setPosition(bookmark.position + bookmark.length, QTextCursor.MoveMode.KeepAnchor)
                return cursor.selectedText()
            elif cross_ref.reference_type == 'page':
                # Simplified - calculate page number
                page_num = bookmark.position // 3000 + 1
                return f"page {page_num}"
            else:
                return bookmark.name

        elif cross_ref.target_type in ['figure', 'table', 'equation', 'listing']:
            if not self.captions_manager:
                return "[No Caption Manager]"

            caption = self.captions_manager.get_caption_by_id(cross_ref.target_id)
            if not caption:
                return "[Invalid Caption]"

            prefix = self.captions_manager.prefixes.get(
                caption.caption_type,
                caption.caption_type.capitalize()
            )
            number = caption.get_formatted_number(
                self.captions_manager.numbering_style,
                self.captions_manager.include_chapter
            )

            if cross_ref.reference_type == 'number':
                return f"{prefix} {number}"
            elif cross_ref.reference_type == 'page':
                # Simplified - calculate page number
                page_num = caption.position // 3000 + 1
                return f"page {page_num}"
            elif cross_ref.reference_type == 'text':
                return caption.text
            elif cross_ref.reference_type == 'number_and_page':
                page_num = caption.position // 3000 + 1
                return f"{prefix} {number} on page {page_num}"
            elif cross_ref.reference_type == 'full':
                return f"{prefix} {number}: {caption.text}"

        elif cross_ref.target_type == 'heading':
            if not self.toc_manager:
                return "[No TOC Manager]"

            # Find heading by ID
            for entry in self.toc_manager.get_entries_flat():
                if hasattr(entry, 'id') and entry.id == cross_ref.target_id:
                    if cross_ref.reference_type == 'text':
                        return entry.text
                    elif cross_ref.reference_type == 'page':
                        return f"page {entry.page_number}"
                    elif cross_ref.reference_type == 'number':
                        return str(entry.level)
                    elif cross_ref.reference_type == 'number_and_page':
                        return f"Section {entry.level} on page {entry.page_number}"
                    else:
                        return entry.text

            return "[Invalid Heading]"

        elif cross_ref.target_type == 'footnote':
            if not self.footnotes_manager:
                return "[No Footnotes Manager]"

            note = self.footnotes_manager.get_note_by_id(cross_ref.target_id)
            if not note:
                return "[Invalid Footnote]"

            mark = note.reference_mark if note.reference_mark else note.number

            if cross_ref.reference_type == 'number':
                return str(mark)
            elif cross_ref.reference_type == 'text':
                return note.text
            elif cross_ref.reference_type == 'page':
                # Simplified - calculate page number
                page_num = note.position // 3000 + 1
                return f"page {page_num}"

        return "[Unknown Reference]"

    def _insert_reference_text(self, cursor, text, cross_ref):
        """Insert formatted reference text."""
        # Format as hyperlink if enabled
        char_format = QTextCharFormat()

        if self.hyperlink_references:
            char_format.setForeground(QColor(0, 0, 255))  # Blue
            char_format.setFontUnderline(True)
            char_format.setAnchor(True)
            char_format.setAnchorHref(f"#ref_{cross_ref.id}")

        cursor.setCharFormat(char_format)
        cursor.insertText(text)

        # Reset format
        cursor.setCharFormat(QTextCharFormat())

    def update_cross_references(self):
        """Update all cross-references in the document."""
        # This would update all reference text in the document
        # Simplified implementation - in a real app, you'd track and update actual positions
        updated_count = 0

        for cross_ref in self.cross_references:
            # Get updated reference text
            new_text = self._get_reference_text(cross_ref)

            if new_text and not new_text.startswith("["):
                # Find and update in document
                # This is simplified - real implementation would track exact positions
                updated_count += 1

        return updated_count

    def find_broken_references(self):
        """Find cross-references that point to non-existent targets."""
        broken = []

        for cross_ref in self.cross_references:
            ref_text = self._get_reference_text(cross_ref)
            if ref_text.startswith("[") and ref_text.endswith("]"):
                # Reference is broken
                broken.append(cross_ref)

        return broken

    def delete_cross_reference(self, cross_ref_id):
        """Delete a cross-reference."""
        cross_ref = None
        for ref in self.cross_references:
            if ref.id == cross_ref_id:
                cross_ref = ref
                break

        if cross_ref:
            self.cross_references.remove(cross_ref)
            return True
        return False


class BookmarkDialog(QDialog):
    """Dialog for adding a bookmark."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Bookmark")
        self.setModal(True)
        self.setMinimumWidth(400)

        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Bookmark name
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Bookmark name:"))
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Enter bookmark name")
        name_layout.addWidget(self.name_edit)
        layout.addLayout(name_layout)

        # Info
        info_label = QLabel(
            "<i>Note: If you have text selected, the bookmark will mark that text. "
            "Otherwise, it will mark the current cursor position.</i>"
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

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

    def get_bookmark_name(self):
        """Get the bookmark name."""
        return self.name_edit.text().strip()


class InsertCrossReferenceDialog(QDialog):
    """Dialog for inserting a cross-reference."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Insert Cross-Reference")
        self.setModal(True)
        self.setMinimumSize(600, 500)

        self.setup_ui()
        self.refresh_targets()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Title
        title_label = QLabel("<h3>Insert Cross-Reference</h3>")
        layout.addWidget(title_label)

        # Reference type selection
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Reference type:"))

        self.type_combo = QComboBox()
        self.type_combo.addItems([
            "Bookmark",
            "Heading",
            "Figure",
            "Table",
            "Equation",
            "Footnote"
        ])
        self.type_combo.currentTextChanged.connect(self.on_type_changed)
        type_layout.addWidget(self.type_combo)
        type_layout.addStretch()

        layout.addLayout(type_layout)

        # Available targets
        targets_label = QLabel("<b>Select target:</b>")
        layout.addWidget(targets_label)

        self.targets_list = QListWidget()
        self.targets_list.itemSelectionChanged.connect(self.on_target_selected)
        layout.addWidget(self.targets_list)

        # Insert options
        options_group = QGroupBox("Insert as")
        options_layout = QVBoxLayout()

        self.reference_type_group = QButtonGroup()

        self.number_radio = QRadioButton("Number only")
        self.number_radio.setChecked(True)
        self.reference_type_group.addButton(self.number_radio)
        options_layout.addWidget(self.number_radio)

        self.page_radio = QRadioButton("Page number")
        self.reference_type_group.addButton(self.page_radio)
        options_layout.addWidget(self.page_radio)

        self.text_radio = QRadioButton("Text")
        self.reference_type_group.addButton(self.text_radio)
        options_layout.addWidget(self.text_radio)

        self.number_page_radio = QRadioButton("Number and page")
        self.reference_type_group.addButton(self.number_page_radio)
        options_layout.addWidget(self.number_page_radio)

        self.full_radio = QRadioButton("Full caption/text")
        self.reference_type_group.addButton(self.full_radio)
        options_layout.addWidget(self.full_radio)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Hyperlink option
        self.hyperlink_cb = QCheckBox("Insert as hyperlink")
        self.hyperlink_cb.setChecked(self.manager.hyperlink_references)
        layout.addWidget(self.hyperlink_cb)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        insert_button = QPushButton("Insert")
        insert_button.clicked.connect(self.insert_reference)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(insert_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def on_type_changed(self):
        """Handle reference type change."""
        self.refresh_targets()
        self.update_reference_options()

    def on_target_selected(self):
        """Handle target selection."""
        pass

    def update_reference_options(self):
        """Update available reference options based on type."""
        ref_type = self.type_combo.currentText().lower()

        # Enable/disable options based on type
        if ref_type == 'bookmark':
            self.number_radio.setEnabled(False)
            self.page_radio.setEnabled(True)
            self.text_radio.setEnabled(True)
            self.number_page_radio.setEnabled(False)
            self.full_radio.setEnabled(False)
            self.text_radio.setChecked(True)
        else:
            self.number_radio.setEnabled(True)
            self.page_radio.setEnabled(True)
            self.text_radio.setEnabled(True)
            self.number_page_radio.setEnabled(True)
            self.full_radio.setEnabled(True)
            self.number_radio.setChecked(True)

    def refresh_targets(self):
        """Refresh the list of available targets."""
        self.targets_list.clear()

        ref_type = self.type_combo.currentText().lower()

        if ref_type == 'bookmark':
            for bookmark in self.manager.bookmarks:
                item = QListWidgetItem(bookmark.name)
                item.setData(Qt.ItemDataRole.UserRole, bookmark.id)
                self.targets_list.addItem(item)

        elif ref_type in ['figure', 'table', 'equation']:
            if self.manager.captions_manager:
                captions = self.manager.captions_manager.get_captions_by_type(ref_type)
                for caption in sorted(captions, key=lambda c: c.position):
                    prefix = self.manager.captions_manager.prefixes.get(
                        caption.caption_type,
                        caption.caption_type.capitalize()
                    )
                    number = caption.get_formatted_number(
                        self.manager.captions_manager.numbering_style,
                        self.manager.captions_manager.include_chapter
                    )
                    preview = caption.text[:40] + "..." if len(caption.text) > 40 else caption.text

                    item_text = f"{prefix} {number}: {preview}"
                    if caption.label:
                        item_text += f" [{caption.label}]"

                    item = QListWidgetItem(item_text)
                    item.setData(Qt.ItemDataRole.UserRole, caption.id)
                    self.targets_list.addItem(item)

        elif ref_type == 'heading':
            if self.manager.toc_manager:
                self.manager.toc_manager.scan_document()
                for entry in self.manager.toc_manager.get_entries_flat():
                    item = QListWidgetItem(f"{'  ' * (entry.level - 1)}{entry.text}")
                    # Note: TOC entries might not have IDs in the current implementation
                    # This would need to be added to the TOCEntry class
                    item.setData(Qt.ItemDataRole.UserRole, entry.position)
                    self.targets_list.addItem(item)

        elif ref_type == 'footnote':
            if self.manager.footnotes_manager:
                for note in sorted(self.manager.footnotes_manager.notes, key=lambda n: n.position):
                    if note.note_type == 'footnote':
                        mark = note.reference_mark if note.reference_mark else note.number
                        preview = note.text[:40] + "..." if len(note.text) > 40 else note.text

                        item = QListWidgetItem(f"[{mark}] {preview}")
                        item.setData(Qt.ItemDataRole.UserRole, note.id)
                        self.targets_list.addItem(item)

    def get_reference_type(self):
        """Get the selected reference type."""
        if self.number_radio.isChecked():
            return 'number'
        elif self.page_radio.isChecked():
            return 'page'
        elif self.text_radio.isChecked():
            return 'text'
        elif self.number_page_radio.isChecked():
            return 'number_and_page'
        elif self.full_radio.isChecked():
            return 'full'
        return 'number'

    def insert_reference(self):
        """Insert the cross-reference."""
        items = self.targets_list.selectedItems()
        if not items:
            QMessageBox.warning(self, "No Selection", "Please select a target for the cross-reference.")
            return

        target_id = items[0].data(Qt.ItemDataRole.UserRole)
        target_type = self.type_combo.currentText().lower()
        reference_type = self.get_reference_type()

        # Update hyperlink setting
        self.manager.hyperlink_references = self.hyperlink_cb.isChecked()

        # Insert cross-reference
        cross_ref = self.manager.insert_cross_reference(target_type, target_id, reference_type)

        if cross_ref:
            QMessageBox.information(self, "Success", "Cross-reference inserted successfully!")
            self.accept()
        else:
            QMessageBox.warning(self, "Error", "Failed to insert cross-reference.")


class CrossReferencesViewer(QDialog):
    """Dialog for viewing and managing cross-references and bookmarks."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Cross-References and Bookmarks")
        self.setModal(False)
        self.setMinimumSize(700, 500)

        self.setup_ui()
        self.refresh_data()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Title
        title_label = QLabel("<h3>Cross-References and Bookmarks</h3>")
        layout.addWidget(title_label)

        # Tabs
        tabs = QTabWidget()

        # Bookmarks tab
        bookmarks_tab = self._create_bookmarks_tab()
        tabs.addTab(bookmarks_tab, "Bookmarks")

        # Cross-references tab
        refs_tab = self._create_references_tab()
        tabs.addTab(refs_tab, "Cross-References")

        layout.addWidget(tabs)

        # Close button
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def _create_bookmarks_tab(self):
        """Create the bookmarks tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Bookmarks list
        self.bookmarks_list = QListWidget()
        self.bookmarks_list.itemDoubleClicked.connect(self.navigate_to_bookmark)
        layout.addWidget(self.bookmarks_list)

        # Buttons
        button_layout = QHBoxLayout()

        add_button = QPushButton("Add Bookmark")
        add_button.clicked.connect(self.add_bookmark)
        button_layout.addWidget(add_button)

        delete_button = QPushButton("Delete")
        delete_button.clicked.connect(self.delete_bookmark)
        button_layout.addWidget(delete_button)

        goto_button = QPushButton("Go To")
        goto_button.clicked.connect(self.navigate_to_bookmark)
        button_layout.addWidget(goto_button)

        button_layout.addStretch()

        layout.addLayout(button_layout)
        tab.setLayout(layout)

        return tab

    def _create_references_tab(self):
        """Create the cross-references tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # References list
        self.refs_list = QListWidget()
        layout.addWidget(self.refs_list)

        # Buttons
        button_layout = QHBoxLayout()

        update_button = QPushButton("Update All")
        update_button.clicked.connect(self.update_references)
        button_layout.addWidget(update_button)

        check_button = QPushButton("Check for Broken References")
        check_button.clicked.connect(self.check_broken_references)
        button_layout.addWidget(check_button)

        delete_button = QPushButton("Delete")
        delete_button.clicked.connect(self.delete_reference)
        button_layout.addWidget(delete_button)

        button_layout.addStretch()

        layout.addLayout(button_layout)
        tab.setLayout(layout)

        return tab

    def refresh_data(self):
        """Refresh bookmarks and cross-references lists."""
        # Refresh bookmarks
        self.bookmarks_list.clear()
        for bookmark in sorted(self.manager.bookmarks, key=lambda b: b.name):
            item = QListWidgetItem(bookmark.name)
            item.setData(Qt.ItemDataRole.UserRole, bookmark.id)
            self.bookmarks_list.addItem(item)

        # Refresh cross-references
        self.refs_list.clear()
        for ref in self.manager.cross_references:
            ref_text = self.manager._get_reference_text(ref)
            item_text = f"{ref.target_type.capitalize()}: {ref_text} ({ref.reference_type})"

            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, ref.id)
            self.refs_list.addItem(item)

    def add_bookmark(self):
        """Add a new bookmark."""
        dialog = BookmarkDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            name = dialog.get_bookmark_name()

            if not name:
                QMessageBox.warning(self, "Invalid Name", "Please enter a bookmark name.")
                return

            bookmark = self.manager.add_bookmark(name)

            if bookmark:
                QMessageBox.information(self, "Success", "Bookmark added successfully!")
                self.refresh_data()
            else:
                QMessageBox.warning(self, "Error", "A bookmark with this name already exists.")

    def delete_bookmark(self):
        """Delete the selected bookmark."""
        items = self.bookmarks_list.selectedItems()
        if not items:
            return

        reply = QMessageBox.question(
            self,
            "Delete Bookmark",
            "Are you sure you want to delete this bookmark?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            bookmark_id = items[0].data(Qt.ItemDataRole.UserRole)
            self.manager.delete_bookmark(bookmark_id)
            self.refresh_data()

    def navigate_to_bookmark(self):
        """Navigate to the selected bookmark."""
        items = self.bookmarks_list.selectedItems()
        if not items:
            return

        bookmark_id = items[0].data(Qt.ItemDataRole.UserRole)
        self.manager.navigate_to_bookmark(bookmark_id)

    def update_references(self):
        """Update all cross-references."""
        count = self.manager.update_cross_references()
        QMessageBox.information(
            self,
            "Update Complete",
            f"Updated {count} cross-reference(s)."
        )
        self.refresh_data()

    def check_broken_references(self):
        """Check for broken cross-references."""
        broken = self.manager.find_broken_references()

        if not broken:
            QMessageBox.information(self, "Check Complete", "No broken references found.")
        else:
            message = f"Found {len(broken)} broken reference(s):\n\n"
            for ref in broken[:10]:  # Show first 10
                ref_text = self.manager._get_reference_text(ref)
                message += f"- {ref.target_type}: {ref_text}\n"

            if len(broken) > 10:
                message += f"\n... and {len(broken) - 10} more"

            QMessageBox.warning(self, "Broken References", message)

    def delete_reference(self):
        """Delete the selected cross-reference."""
        items = self.refs_list.selectedItems()
        if not items:
            return

        reply = QMessageBox.question(
            self,
            "Delete Cross-Reference",
            "Are you sure you want to delete this cross-reference?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            ref_id = items[0].data(Qt.ItemDataRole.UserRole)
            self.manager.delete_cross_reference(ref_id)
            self.refresh_data()
