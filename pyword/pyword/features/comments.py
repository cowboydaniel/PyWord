"""
Comments System for PyWord.

This module provides a comprehensive commenting system for collaborative document editing,
including thread support, mentions, and comment resolution.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit,
                               QPushButton, QListWidget, QListWidgetItem, QInputDialog,
                               QMessageBox, QWidget, QSplitter, QFrame, QLineEdit,
                               QComboBox, QCheckBox, QGroupBox, QScrollArea)
from PySide6.QtCore import Qt, Signal, QDateTime, QPoint
from PySide6.QtGui import QTextCursor, QTextCharFormat, QColor, QFont
from datetime import datetime
import json
import uuid


class Comment:
    """Represents a comment in the document."""

    def __init__(self, text, author, position_start, position_end,
                 parent_id=None, timestamp=None):
        self.id = str(uuid.uuid4())
        self.text = text
        self.author = author
        self.position_start = position_start
        self.position_end = position_end
        self.timestamp = timestamp or datetime.now()
        self.parent_id = parent_id  # For threaded comments/replies
        self.resolved = False
        self.edited = False
        self.edit_timestamp = None
        self.replies = []  # List of reply comment IDs
        self.mentions = []  # List of mentioned users
        self.tags = []  # Tags for categorization

    def to_dict(self):
        """Convert comment to dictionary for serialization."""
        return {
            'id': self.id,
            'text': self.text,
            'author': self.author,
            'position_start': self.position_start,
            'position_end': self.position_end,
            'timestamp': self.timestamp.isoformat(),
            'parent_id': self.parent_id,
            'resolved': self.resolved,
            'edited': self.edited,
            'edit_timestamp': self.edit_timestamp.isoformat() if self.edit_timestamp else None,
            'replies': self.replies,
            'mentions': self.mentions,
            'tags': self.tags
        }

    @staticmethod
    def from_dict(data):
        """Create Comment from dictionary."""
        comment = Comment(
            data['text'],
            data['author'],
            data['position_start'],
            data['position_end'],
            data.get('parent_id'),
            datetime.fromisoformat(data['timestamp'])
        )
        comment.id = data['id']
        comment.resolved = data.get('resolved', False)
        comment.edited = data.get('edited', False)
        if data.get('edit_timestamp'):
            comment.edit_timestamp = datetime.fromisoformat(data['edit_timestamp'])
        comment.replies = data.get('replies', [])
        comment.mentions = data.get('mentions', [])
        comment.tags = data.get('tags', [])
        return comment

    def get_selected_text(self, document):
        """Get the text that this comment refers to."""
        cursor = QTextCursor(document)
        cursor.setPosition(self.position_start)
        cursor.setPosition(self.position_end, QTextCursor.MoveMode.KeepAnchor)
        return cursor.selectedText()


class CommentsManager:
    """Manages comments in a document."""

    def __init__(self, parent):
        self.parent = parent
        self.comments = []
        self.current_author = "Unknown User"
        self.show_comments = True

        # Formatting for comment highlighting
        self.comment_format = QTextCharFormat()
        self.comment_format.setBackground(QColor(255, 255, 200))  # Light yellow
        self.comment_format.setProperty(QTextCharFormat.Property.UserProperty, "comment")

    def set_author(self, author_name):
        """Set the current author name."""
        self.current_author = author_name

    def add_comment(self, text, position_start=None, position_end=None, parent_id=None):
        """Add a new comment."""
        cursor = self.parent.textCursor()

        # If positions not specified, use current selection
        if position_start is None:
            if cursor.hasSelection():
                position_start = cursor.selectionStart()
                position_end = cursor.selectionEnd()
            else:
                position_start = cursor.position()
                position_end = cursor.position()

        # Extract mentions from text (words starting with @)
        mentions = [word[1:] for word in text.split() if word.startswith('@')]

        comment = Comment(
            text,
            self.current_author,
            position_start,
            position_end,
            parent_id
        )
        comment.mentions = mentions

        self.comments.append(comment)

        # If this is a reply, add to parent's replies list
        if parent_id:
            parent = self.get_comment_by_id(parent_id)
            if parent:
                parent.replies.append(comment.id)

        # Highlight the commented text
        if self.show_comments and position_start != position_end:
            self._highlight_comment(comment)

        return comment

    def edit_comment(self, comment_id, new_text):
        """Edit an existing comment."""
        comment = self.get_comment_by_id(comment_id)
        if comment:
            comment.text = new_text
            comment.edited = True
            comment.edit_timestamp = datetime.now()

            # Update mentions
            comment.mentions = [word[1:] for word in new_text.split() if word.startswith('@')]

            return True
        return False

    def delete_comment(self, comment_id):
        """Delete a comment and its replies."""
        comment = self.get_comment_by_id(comment_id)
        if not comment:
            return False

        # Delete all replies first
        for reply_id in comment.replies[:]:  # Use slice to avoid modifying list during iteration
            self.delete_comment(reply_id)

        # Remove highlighting
        if comment.position_start != comment.position_end:
            self._remove_highlight(comment)

        # Remove from parent's reply list if it's a reply
        if comment.parent_id:
            parent = self.get_comment_by_id(comment.parent_id)
            if parent and comment.id in parent.replies:
                parent.replies.remove(comment.id)

        # Remove comment
        self.comments.remove(comment)
        return True

    def resolve_comment(self, comment_id):
        """Mark a comment as resolved."""
        comment = self.get_comment_by_id(comment_id)
        if comment:
            comment.resolved = True

            # Resolve all replies as well
            for reply_id in comment.replies:
                reply = self.get_comment_by_id(reply_id)
                if reply:
                    reply.resolved = True

            return True
        return False

    def unresolve_comment(self, comment_id):
        """Mark a comment as unresolved."""
        comment = self.get_comment_by_id(comment_id)
        if comment:
            comment.resolved = False
            return True
        return False

    def get_comment_by_id(self, comment_id):
        """Get a comment by its ID."""
        for comment in self.comments:
            if comment.id == comment_id:
                return comment
        return None

    def get_comments_by_author(self, author):
        """Get all comments by a specific author."""
        return [c for c in self.comments if c.author == author]

    def get_active_comments(self):
        """Get all unresolved comments."""
        return [c for c in self.comments if not c.resolved and not c.parent_id]

    def get_resolved_comments(self):
        """Get all resolved comments."""
        return [c for c in self.comments if c.resolved and not c.parent_id]

    def get_all_root_comments(self):
        """Get all root-level comments (not replies)."""
        return [c for c in self.comments if not c.parent_id]

    def get_replies(self, comment_id):
        """Get all direct replies to a comment."""
        comment = self.get_comment_by_id(comment_id)
        if comment:
            return [self.get_comment_by_id(reply_id) for reply_id in comment.replies]
        return []

    def get_comment_thread(self, comment_id):
        """Get a complete comment thread (comment and all its replies)."""
        comment = self.get_comment_by_id(comment_id)
        if not comment:
            return []

        thread = [comment]
        for reply_id in comment.replies:
            reply = self.get_comment_by_id(reply_id)
            if reply:
                thread.append(reply)
                # Recursively get nested replies
                thread.extend(self.get_comment_thread(reply_id)[1:])

        return thread

    def get_mentions(self, author):
        """Get all comments that mention a specific user."""
        return [c for c in self.comments if author in c.mentions]

    def add_tag(self, comment_id, tag):
        """Add a tag to a comment."""
        comment = self.get_comment_by_id(comment_id)
        if comment and tag not in comment.tags:
            comment.tags.append(tag)
            return True
        return False

    def remove_tag(self, comment_id, tag):
        """Remove a tag from a comment."""
        comment = self.get_comment_by_id(comment_id)
        if comment and tag in comment.tags:
            comment.tags.remove(tag)
            return True
        return False

    def get_comments_by_tag(self, tag):
        """Get all comments with a specific tag."""
        return [c for c in self.comments if tag in c.tags]

    def navigate_to_comment(self, comment_id):
        """Navigate to the position of a comment in the document."""
        comment = self.get_comment_by_id(comment_id)
        if comment:
            cursor = self.parent.textCursor()
            cursor.setPosition(comment.position_start)
            self.parent.setTextCursor(cursor)
            self.parent.ensureCursorVisible()
            return True
        return False

    def toggle_show_comments(self):
        """Toggle showing/hiding comment highlights."""
        self.show_comments = not self.show_comments

        if self.show_comments:
            # Reapply all highlights
            for comment in self.comments:
                if comment.position_start != comment.position_end:
                    self._highlight_comment(comment)
        else:
            # Remove all highlights
            for comment in self.comments:
                if comment.position_start != comment.position_end:
                    self._remove_highlight(comment)

    def _highlight_comment(self, comment):
        """Apply highlighting to commented text."""
        cursor = QTextCursor(self.parent.document())
        cursor.setPosition(comment.position_start)
        cursor.setPosition(comment.position_end, QTextCursor.MoveMode.KeepAnchor)

        format = QTextCharFormat()
        if comment.resolved:
            format.setBackground(QColor(200, 255, 200))  # Light green
        else:
            format.setBackground(QColor(255, 255, 200))  # Light yellow

        cursor.mergeCharFormat(format)

    def _remove_highlight(self, comment):
        """Remove highlighting from commented text."""
        cursor = QTextCursor(self.parent.document())
        cursor.setPosition(comment.position_start)
        cursor.setPosition(comment.position_end, QTextCursor.MoveMode.KeepAnchor)

        # Reset to default formatting
        format = QTextCharFormat()
        format.setBackground(QColor(Qt.GlobalColor.white))
        cursor.setCharFormat(format)

    def export_comments(self, file_path):
        """Export comments to a JSON file."""
        try:
            data = {
                'comments': [c.to_dict() for c in self.comments],
                'author': self.current_author,
                'exported': datetime.now().isoformat()
            }

            with open(file_path, 'w') as f:
                json.dump(data, f, indent=2)

            return True
        except Exception as e:
            print(f"Error exporting comments: {e}")
            return False

    def import_comments(self, file_path):
        """Import comments from a JSON file."""
        try:
            with open(file_path, 'r') as f:
                data = json.load(f)

            for comment_data in data.get('comments', []):
                comment = Comment.from_dict(comment_data)
                self.comments.append(comment)

                # Apply highlighting if enabled
                if self.show_comments and comment.position_start != comment.position_end:
                    self._highlight_comment(comment)

            return True
        except Exception as e:
            print(f"Error importing comments: {e}")
            return False


class CommentDialog(QDialog):
    """Dialog for adding or editing a comment."""

    def __init__(self, comment=None, parent=None):
        super().__init__(parent)
        self.comment = comment
        self.setWindowTitle("Edit Comment" if comment else "Add Comment")
        self.setModal(True)
        self.setMinimumWidth(400)

        self.setup_ui()

        # Load existing comment data
        if comment:
            self.text_edit.setPlainText(comment.text)
            for tag in comment.tags:
                self.tags_edit.setText(", ".join(comment.tags))

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Comment text
        text_label = QLabel("Comment:")
        layout.addWidget(text_label)

        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Enter your comment here... (Use @ to mention someone)")
        self.text_edit.setMinimumHeight(100)
        layout.addWidget(self.text_edit)

        # Tags
        tags_label = QLabel("Tags (comma-separated):")
        layout.addWidget(tags_label)

        self.tags_edit = QLineEdit()
        self.tags_edit.setPlaceholderText("e.g., important, question, todo")
        layout.addWidget(self.tags_edit)

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

    def get_comment_text(self):
        """Get the comment text."""
        return self.text_edit.toPlainText()

    def get_tags(self):
        """Get the list of tags."""
        tags_text = self.tags_edit.text().strip()
        if tags_text:
            return [tag.strip() for tag in tags_text.split(',') if tag.strip()]
        return []


class CommentsViewer(QDialog):
    """Main dialog for viewing and managing comments."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Comments")
        self.setModal(False)
        self.setMinimumSize(700, 600)

        self.setup_ui()
        self.refresh_comments()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Header with controls
        header_layout = QHBoxLayout()

        header_label = QLabel("<h3>Document Comments</h3>")
        header_layout.addWidget(header_label)

        header_layout.addStretch()

        # Filter
        filter_label = QLabel("Show:")
        header_layout.addWidget(filter_label)

        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["All Comments", "Active", "Resolved", "My Comments", "Mentions"])
        self.filter_combo.currentTextChanged.connect(self.refresh_comments)
        header_layout.addWidget(self.filter_combo)

        # Show/hide highlights
        self.show_highlights_checkbox = QCheckBox("Show Highlights")
        self.show_highlights_checkbox.setChecked(True)
        self.show_highlights_checkbox.toggled.connect(self.manager.toggle_show_comments)
        header_layout.addWidget(self.show_highlights_checkbox)

        layout.addLayout(header_layout)

        # Splitter for comments list and details
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Comments list
        self.comments_list = QListWidget()
        self.comments_list.itemSelectionChanged.connect(self.on_comment_selected)
        self.comments_list.itemDoubleClicked.connect(self.on_comment_double_clicked)
        splitter.addWidget(self.comments_list)

        # Comment details panel
        details_panel = QWidget()
        details_layout = QVBoxLayout(details_panel)

        self.details_text = QTextEdit()
        self.details_text.setReadOnly(True)
        details_layout.addWidget(self.details_text)

        # Replies section
        replies_label = QLabel("<b>Replies:</b>")
        details_layout.addWidget(replies_label)

        self.replies_list = QListWidget()
        self.replies_list.setMaximumHeight(150)
        details_layout.addWidget(self.replies_list)

        # Action buttons
        button_layout = QHBoxLayout()

        self.reply_button = QPushButton("Reply")
        self.reply_button.clicked.connect(self.reply_to_comment)
        button_layout.addWidget(self.reply_button)

        self.edit_button = QPushButton("Edit")
        self.edit_button.clicked.connect(self.edit_comment)
        button_layout.addWidget(self.edit_button)

        self.delete_button = QPushButton("Delete")
        self.delete_button.clicked.connect(self.delete_comment)
        button_layout.addWidget(self.delete_button)

        self.resolve_button = QPushButton("Resolve")
        self.resolve_button.clicked.connect(self.resolve_comment)
        button_layout.addWidget(self.resolve_button)

        self.navigate_button = QPushButton("Go To")
        self.navigate_button.clicked.connect(self.navigate_to_comment)
        button_layout.addWidget(self.navigate_button)

        details_layout.addLayout(button_layout)

        splitter.addWidget(details_panel)
        splitter.setSizes([300, 400])

        layout.addWidget(splitter)

        # Bottom buttons
        bottom_layout = QHBoxLayout()

        new_button = QPushButton("New Comment")
        new_button.clicked.connect(self.new_comment)
        bottom_layout.addWidget(new_button)

        bottom_layout.addStretch()

        export_button = QPushButton("Export")
        export_button.clicked.connect(self.export_comments)
        bottom_layout.addWidget(export_button)

        import_button = QPushButton("Import")
        import_button.clicked.connect(self.import_comments)
        bottom_layout.addWidget(import_button)

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        bottom_layout.addWidget(close_button)

        layout.addLayout(bottom_layout)

        self.setLayout(layout)

    def refresh_comments(self):
        """Refresh the comments list."""
        self.comments_list.clear()

        # Get filtered comments
        filter_text = self.filter_combo.currentText()

        if filter_text == "All Comments":
            comments = self.manager.get_all_root_comments()
        elif filter_text == "Active":
            comments = self.manager.get_active_comments()
        elif filter_text == "Resolved":
            comments = self.manager.get_resolved_comments()
        elif filter_text == "My Comments":
            comments = [c for c in self.manager.get_comments_by_author(self.manager.current_author)
                       if not c.parent_id]
        elif filter_text == "Mentions":
            comments = [c for c in self.manager.get_mentions(self.manager.current_author)
                       if not c.parent_id]
        else:
            comments = self.manager.get_all_root_comments()

        # Add to list
        for comment in comments:
            preview = comment.text[:50] + "..." if len(comment.text) > 50 else comment.text
            status = " [RESOLVED]" if comment.resolved else ""
            item_text = f"{comment.author} - {preview}{status}"

            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, comment.id)
            self.comments_list.addItem(item)

    def on_comment_selected(self):
        """Handle comment selection."""
        items = self.comments_list.selectedItems()
        if not items:
            self.details_text.clear()
            self.replies_list.clear()
            return

        comment_id = items[0].data(Qt.ItemDataRole.UserRole)
        comment = self.manager.get_comment_by_id(comment_id)

        if comment:
            # Show comment details
            details_html = f"<b>Author:</b> {comment.author}<br>"
            details_html += f"<b>Date:</b> {comment.timestamp.strftime('%Y-%m-%d %H:%M:%S')}<br>"

            if comment.edited:
                details_html += f"<b>Edited:</b> {comment.edit_timestamp.strftime('%Y-%m-%d %H:%M:%S')}<br>"

            if comment.tags:
                details_html += f"<b>Tags:</b> {', '.join(comment.tags)}<br>"

            if comment.mentions:
                details_html += f"<b>Mentions:</b> @{', @'.join(comment.mentions)}<br>"

            details_html += f"<br><b>Comment:</b><br>{comment.text}<br>"

            # Show referenced text
            if comment.position_start != comment.position_end:
                referenced_text = comment.get_selected_text(self.manager.parent.document())
                details_html += f"<br><b>Referenced text:</b><br><i>{referenced_text}</i>"

            self.details_text.setHtml(details_html)

            # Show replies
            self.replies_list.clear()
            replies = self.manager.get_replies(comment_id)
            for reply in replies:
                reply_text = f"{reply.author}: {reply.text[:50]}"
                if len(reply.text) > 50:
                    reply_text += "..."
                self.replies_list.addItem(reply_text)

            # Update button states
            self.resolve_button.setText("Unresolve" if comment.resolved else "Resolve")

    def on_comment_double_clicked(self, item):
        """Handle double-click on comment."""
        self.navigate_to_comment()

    def new_comment(self):
        """Create a new comment."""
        dialog = CommentDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            text = dialog.get_comment_text()
            tags = dialog.get_tags()

            if text:
                comment = self.manager.add_comment(text)
                for tag in tags:
                    self.manager.add_tag(comment.id, tag)

                self.refresh_comments()

    def reply_to_comment(self):
        """Reply to the selected comment."""
        items = self.comments_list.selectedItems()
        if not items:
            return

        comment_id = items[0].data(Qt.ItemDataRole.UserRole)

        dialog = CommentDialog(parent=self)
        dialog.setWindowTitle("Reply to Comment")

        if dialog.exec() == QDialog.DialogCode.Accepted:
            text = dialog.get_comment_text()
            if text:
                self.manager.add_comment(text, parent_id=comment_id)
                self.refresh_comments()
                # Reselect the parent comment to show updated replies
                self.on_comment_selected()

    def edit_comment(self):
        """Edit the selected comment."""
        items = self.comments_list.selectedItems()
        if not items:
            return

        comment_id = items[0].data(Qt.ItemDataRole.UserRole)
        comment = self.manager.get_comment_by_id(comment_id)

        if comment:
            dialog = CommentDialog(comment, parent=self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                new_text = dialog.get_comment_text()
                self.manager.edit_comment(comment_id, new_text)

                # Update tags
                new_tags = dialog.get_tags()
                comment.tags = new_tags

                self.refresh_comments()
                self.on_comment_selected()

    def delete_comment(self):
        """Delete the selected comment."""
        items = self.comments_list.selectedItems()
        if not items:
            return

        reply = QMessageBox.question(
            self,
            "Delete Comment",
            "Are you sure you want to delete this comment and all its replies?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            comment_id = items[0].data(Qt.ItemDataRole.UserRole)
            self.manager.delete_comment(comment_id)
            self.refresh_comments()

    def resolve_comment(self):
        """Resolve or unresolve the selected comment."""
        items = self.comments_list.selectedItems()
        if not items:
            return

        comment_id = items[0].data(Qt.ItemDataRole.UserRole)
        comment = self.manager.get_comment_by_id(comment_id)

        if comment:
            if comment.resolved:
                self.manager.unresolve_comment(comment_id)
            else:
                self.manager.resolve_comment(comment_id)

            self.refresh_comments()
            self.on_comment_selected()

    def navigate_to_comment(self):
        """Navigate to the selected comment in the document."""
        items = self.comments_list.selectedItems()
        if not items:
            return

        comment_id = items[0].data(Qt.ItemDataRole.UserRole)
        self.manager.navigate_to_comment(comment_id)

    def export_comments(self):
        """Export comments to a file."""
        from PySide6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Comments",
            "",
            "JSON Files (*.json)"
        )

        if file_path:
            if self.manager.export_comments(file_path):
                QMessageBox.information(self, "Success", "Comments exported successfully!")
            else:
                QMessageBox.warning(self, "Error", "Failed to export comments.")

    def import_comments(self):
        """Import comments from a file."""
        from PySide6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Import Comments",
            "",
            "JSON Files (*.json)"
        )

        if file_path:
            if self.manager.import_comments(file_path):
                QMessageBox.information(self, "Success", "Comments imported successfully!")
                self.refresh_comments()
            else:
                QMessageBox.warning(self, "Error", "Failed to import comments.")
