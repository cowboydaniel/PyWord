"""
Track Changes for PyWord.

This module provides track changes functionality for collaborative document editing,
allowing users to see who made what changes and when.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
                               QPushButton, QGroupBox, QListWidget, QListWidgetItem,
                               QTextEdit, QCheckBox, QSplitter, QWidget, QMessageBox,
                               QInputDialog, QFrame)
from PySide6.QtCore import Qt, QDateTime, Signal
from PySide6.QtGui import QColor, QTextCharFormat, QTextCursor, QFont
from datetime import datetime
import json


class ChangeType:
    """Types of changes that can be tracked."""
    INSERTION = "Insertion"
    DELETION = "Deletion"
    FORMATTING = "Formatting"
    MOVE = "Move"


class Change:
    """Represents a tracked change in the document."""

    def __init__(self, change_type, position, content, author, timestamp=None):
        self.change_type = change_type
        self.position = position
        self.content = content
        self.author = author
        self.timestamp = timestamp or datetime.now()
        self.accepted = False
        self.rejected = False
        self.id = f"{change_type}_{position}_{self.timestamp.timestamp()}"

    def to_dict(self):
        """Convert change to dictionary for serialization."""
        return {
            'change_type': self.change_type,
            'position': self.position,
            'content': self.content,
            'author': self.author,
            'timestamp': self.timestamp.isoformat(),
            'accepted': self.accepted,
            'rejected': self.rejected,
            'id': self.id
        }

    @staticmethod
    def from_dict(data):
        """Create Change from dictionary."""
        change = Change(
            data['change_type'],
            data['position'],
            data['content'],
            data['author'],
            datetime.fromisoformat(data['timestamp'])
        )
        change.accepted = data.get('accepted', False)
        change.rejected = data.get('rejected', False)
        change.id = data.get('id', change.id)
        return change


class TrackChangesManager:
    """Manages track changes functionality."""

    def __init__(self, parent):
        self.parent = parent
        self.enabled = False
        self.changes = []
        self.current_author = "Unknown User"
        self.show_changes = True

        # Formatting for different change types
        self.insertion_format = QTextCharFormat()
        self.insertion_format.setForeground(QColor(0, 128, 0))  # Green
        self.insertion_format.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SingleUnderline)

        self.deletion_format = QTextCharFormat()
        self.deletion_format.setForeground(QColor(255, 0, 0))  # Red
        self.deletion_format.setFontStrikeOut(True)

        self.formatting_format = QTextCharFormat()
        self.formatting_format.setBackground(QColor(255, 255, 0, 100))  # Light yellow

    def enable_tracking(self, author_name=None):
        """Enable track changes."""
        if author_name:
            self.current_author = author_name
        elif self.current_author == "Unknown User":
            # Prompt for author name
            self.current_author = self._get_author_name()

        self.enabled = True

        # Connect to text editor signals
        if hasattr(self.parent, 'textChanged'):
            self.parent.textChanged.connect(self._on_text_changed)

    def disable_tracking(self):
        """Disable track changes."""
        self.enabled = False

        # Disconnect signals
        if hasattr(self.parent, 'textChanged'):
            try:
                self.parent.textChanged.disconnect(self._on_text_changed)
            except:
                pass

    def toggle_tracking(self):
        """Toggle track changes on/off."""
        if self.enabled:
            self.disable_tracking()
        else:
            self.enable_tracking()

    def _get_author_name(self):
        """Prompt user for author name."""
        name, ok = QInputDialog.getText(
            self.parent,
            "Author Name",
            "Enter your name for tracking changes:"
        )
        return name if ok and name else "Unknown User"

    def _on_text_changed(self):
        """Handle text changes in the document."""
        if not self.enabled:
            return

        cursor = self.parent.textCursor()

        # Check if this is an insertion or deletion
        if cursor.hasSelection():
            # Text is being replaced/deleted
            self._record_deletion(cursor)
        else:
            # Text is being inserted
            self._record_insertion(cursor)

    def _record_insertion(self, cursor):
        """Record an insertion."""
        position = cursor.position()
        # Get the last character typed
        cursor.movePosition(QTextCursor.MoveOperation.Left, QTextCursor.MoveMode.KeepAnchor, 1)
        content = cursor.selectedText()

        if content:
            change = Change(
                ChangeType.INSERTION,
                position,
                content,
                self.current_author
            )
            self.changes.append(change)

            # Apply formatting to show the change
            if self.show_changes:
                cursor.mergeCharFormat(self.insertion_format)

    def _record_deletion(self, cursor):
        """Record a deletion."""
        position = cursor.selectionStart()
        content = cursor.selectedText()

        change = Change(
            ChangeType.DELETION,
            position,
            content,
            self.current_author
        )
        self.changes.append(change)

        # Apply formatting to show the deletion
        if self.show_changes:
            cursor.mergeCharFormat(self.deletion_format)

    def accept_change(self, change):
        """Accept a change."""
        if change in self.changes:
            change.accepted = True

            # Remove formatting
            if change.change_type == ChangeType.DELETION:
                # Actually delete the text
                cursor = self.parent.textCursor()
                cursor.setPosition(change.position)
                cursor.movePosition(
                    QTextCursor.MoveOperation.Right,
                    QTextCursor.MoveMode.KeepAnchor,
                    len(change.content)
                )
                cursor.removeSelectedText()
            elif change.change_type == ChangeType.INSERTION:
                # Keep the text but remove formatting
                cursor = self.parent.textCursor()
                cursor.setPosition(change.position)
                cursor.movePosition(
                    QTextCursor.MoveOperation.Right,
                    QTextCursor.MoveMode.KeepAnchor,
                    len(change.content)
                )
                # Reset to normal formatting
                normal_format = QTextCharFormat()
                cursor.setCharFormat(normal_format)

            return True
        return False

    def reject_change(self, change):
        """Reject a change."""
        if change in self.changes:
            change.rejected = True

            # Revert the change
            if change.change_type == ChangeType.DELETION:
                # Restore the deleted text
                cursor = self.parent.textCursor()
                cursor.setPosition(change.position)
                cursor.insertText(change.content)
            elif change.change_type == ChangeType.INSERTION:
                # Remove the inserted text
                cursor = self.parent.textCursor()
                cursor.setPosition(change.position)
                cursor.movePosition(
                    QTextCursor.MoveOperation.Right,
                    QTextCursor.MoveMode.KeepAnchor,
                    len(change.content)
                )
                cursor.removeSelectedText()

            return True
        return False

    def accept_all_changes(self):
        """Accept all changes."""
        for change in self.changes:
            if not change.accepted and not change.rejected:
                self.accept_change(change)

    def reject_all_changes(self):
        """Reject all changes."""
        for change in self.changes:
            if not change.accepted and not change.rejected:
                self.reject_change(change)

    def get_changes_by_author(self, author):
        """Get all changes by a specific author."""
        return [c for c in self.changes if c.author == author]

    def get_pending_changes(self):
        """Get all pending (not accepted or rejected) changes."""
        return [c for c in self.changes if not c.accepted and not c.rejected]

    def toggle_show_changes(self):
        """Toggle showing/hiding changes."""
        self.show_changes = not self.show_changes

        if self.show_changes:
            self._apply_change_formatting()
        else:
            self._remove_change_formatting()

    def _apply_change_formatting(self):
        """Apply formatting to show all changes."""
        for change in self.get_pending_changes():
            cursor = self.parent.textCursor()
            cursor.setPosition(change.position)
            cursor.movePosition(
                QTextCursor.MoveOperation.Right,
                QTextCursor.MoveMode.KeepAnchor,
                len(change.content)
            )

            if change.change_type == ChangeType.INSERTION:
                cursor.mergeCharFormat(self.insertion_format)
            elif change.change_type == ChangeType.DELETION:
                cursor.mergeCharFormat(self.deletion_format)

    def _remove_change_formatting(self):
        """Remove formatting from all changes."""
        for change in self.get_pending_changes():
            cursor = self.parent.textCursor()
            cursor.setPosition(change.position)
            cursor.movePosition(
                QTextCursor.MoveOperation.Right,
                QTextCursor.MoveMode.KeepAnchor,
                len(change.content)
            )

            # Reset to normal formatting
            normal_format = QTextCharFormat()
            cursor.setCharFormat(normal_format)

    def export_changes(self, file_path):
        """Export changes to a JSON file."""
        try:
            data = {
                'changes': [c.to_dict() for c in self.changes],
                'author': self.current_author,
                'exported': datetime.now().isoformat()
            }

            with open(file_path, 'w') as f:
                json.dump(data, f, indent=2)

            return True
        except Exception as e:
            print(f"Error exporting changes: {e}")
            return False

    def import_changes(self, file_path):
        """Import changes from a JSON file."""
        try:
            with open(file_path, 'r') as f:
                data = json.load(f)

            for change_data in data.get('changes', []):
                change = Change.from_dict(change_data)
                self.changes.append(change)

            return True
        except Exception as e:
            print(f"Error importing changes: {e}")
            return False


class TrackChangesPanel(QWidget):
    """Panel for reviewing tracked changes."""

    change_accepted = Signal(Change)
    change_rejected = Signal(Change)

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setup_ui()

    def setup_ui(self):
        """Setup the panel UI."""
        layout = QVBoxLayout()

        # Header
        header = QLabel("<b>Track Changes</b>")
        layout.addWidget(header)

        # Controls
        control_layout = QHBoxLayout()

        self.toggle_button = QPushButton("Enable Tracking")
        self.toggle_button.clicked.connect(self.toggle_tracking)
        control_layout.addWidget(self.toggle_button)

        self.show_changes_checkbox = QCheckBox("Show Changes")
        self.show_changes_checkbox.setChecked(True)
        self.show_changes_checkbox.toggled.connect(self.manager.toggle_show_changes)
        control_layout.addWidget(self.show_changes_checkbox)

        layout.addLayout(control_layout)

        # Changes list
        list_label = QLabel("Pending Changes:")
        layout.addWidget(list_label)

        self.changes_list = QListWidget()
        self.changes_list.itemSelectionChanged.connect(self.on_change_selected)
        layout.addWidget(self.changes_list)

        # Change details
        details_label = QLabel("Change Details:")
        layout.addWidget(details_label)

        self.details_text = QTextEdit()
        self.details_text.setReadOnly(True)
        self.details_text.setMaximumHeight(100)
        layout.addWidget(self.details_text)

        # Action buttons
        action_layout = QHBoxLayout()

        accept_button = QPushButton("Accept")
        accept_button.clicked.connect(self.accept_selected)
        action_layout.addWidget(accept_button)

        reject_button = QPushButton("Reject")
        reject_button.clicked.connect(self.reject_selected)
        action_layout.addWidget(reject_button)

        layout.addLayout(action_layout)

        # Batch actions
        batch_layout = QHBoxLayout()

        accept_all_button = QPushButton("Accept All")
        accept_all_button.clicked.connect(self.accept_all)
        batch_layout.addWidget(accept_all_button)

        reject_all_button = QPushButton("Reject All")
        reject_all_button.clicked.connect(self.reject_all)
        batch_layout.addWidget(reject_all_button)

        layout.addLayout(batch_layout)

        self.setLayout(layout)

        # Initial refresh
        self.refresh_changes()

    def toggle_tracking(self):
        """Toggle track changes."""
        self.manager.toggle_tracking()

        if self.manager.enabled:
            self.toggle_button.setText("Disable Tracking")
        else:
            self.toggle_button.setText("Enable Tracking")

    def refresh_changes(self):
        """Refresh the changes list."""
        self.changes_list.clear()

        for change in self.manager.get_pending_changes():
            item_text = f"{change.change_type} by {change.author} at {change.timestamp.strftime('%Y-%m-%d %H:%M')}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, change)
            self.changes_list.addItem(item)

    def on_change_selected(self):
        """Handle change selection."""
        items = self.changes_list.selectedItems()
        if items:
            change = items[0].data(Qt.ItemDataRole.UserRole)
            self.show_change_details(change)

    def show_change_details(self, change):
        """Show details of a change."""
        details = f"<b>Type:</b> {change.change_type}<br>"
        details += f"<b>Author:</b> {change.author}<br>"
        details += f"<b>Time:</b> {change.timestamp.strftime('%Y-%m-%d %H:%M:%S')}<br>"
        details += f"<b>Position:</b> {change.position}<br>"
        details += f"<b>Content:</b> {change.content}"

        self.details_text.setHtml(details)

    def accept_selected(self):
        """Accept the selected change."""
        items = self.changes_list.selectedItems()
        if items:
            change = items[0].data(Qt.ItemDataRole.UserRole)
            self.manager.accept_change(change)
            self.change_accepted.emit(change)
            self.refresh_changes()

    def reject_selected(self):
        """Reject the selected change."""
        items = self.changes_list.selectedItems()
        if items:
            change = items[0].data(Qt.ItemDataRole.UserRole)
            self.manager.reject_change(change)
            self.change_rejected.emit(change)
            self.refresh_changes()

    def accept_all(self):
        """Accept all changes."""
        reply = QMessageBox.question(
            self,
            "Accept All Changes",
            "Are you sure you want to accept all pending changes?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.manager.accept_all_changes()
            self.refresh_changes()

    def reject_all(self):
        """Reject all changes."""
        reply = QMessageBox.question(
            self,
            "Reject All Changes",
            "Are you sure you want to reject all pending changes?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.manager.reject_all_changes()
            self.refresh_changes()


class TrackChangesDialog(QDialog):
    """Dialog for managing tracked changes."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Track Changes")
        self.setModal(False)
        self.setMinimumSize(600, 500)

        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Create the track changes panel
        self.panel = TrackChangesPanel(self.manager)
        layout.addWidget(self.panel)

        # Close button
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)
