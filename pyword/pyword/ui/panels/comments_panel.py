"""Comments Panel for PyWord."""

from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QListWidget,
                             QListWidgetItem, QLabel, QPushButton, QTextEdit,
                             QFrame, QSizePolicy, QMenu)
from PySide6.QtCore import Qt, Signal, QDateTime
from PySide6.QtGui import QFont, QAction


class CommentWidget(QFrame):
    """Widget representing a single comment."""

    # Signals
    reply_requested = Signal(object)  # Emits the comment object
    edit_requested = Signal(object)
    delete_requested = Signal(object)
    resolve_requested = Signal(object)

    def __init__(self, comment, parent=None):
        super().__init__(parent)
        self.comment = comment
        self.setup_ui()

    def setup_ui(self):
        """Setup the comment widget UI."""
        self.setFrameStyle(QFrame.Box | QFrame.Plain)
        self.setStyleSheet("QFrame { border: 1px solid #ccc; border-radius: 3px; padding: 5px; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(5)

        # Header with author and date
        header_layout = QHBoxLayout()

        author_label = QLabel(self.comment.get('author', 'Unknown'))
        author_font = QFont()
        author_font.setBold(True)
        author_label.setFont(author_font)
        header_layout.addWidget(author_label)

        header_layout.addStretch()

        date_label = QLabel(self.comment.get('date', ''))
        date_label.setStyleSheet("color: gray; font-size: 9pt;")
        header_layout.addWidget(date_label)

        layout.addLayout(header_layout)

        # Comment text
        text_label = QLabel(self.comment.get('text', ''))
        text_label.setWordWrap(True)
        text_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(text_label)

        # Reference/context
        if 'context' in self.comment:
            context_label = QLabel(f"Re: \"{self.comment['context']}\"")
            context_label.setStyleSheet("color: gray; font-style: italic; font-size: 9pt;")
            context_label.setWordWrap(True)
            layout.addWidget(context_label)

        # Action buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        reply_btn = QPushButton("Reply")
        reply_btn.setFlat(True)
        reply_btn.clicked.connect(lambda: self.reply_requested.emit(self.comment))
        button_layout.addWidget(reply_btn)

        if not self.comment.get('resolved', False):
            resolve_btn = QPushButton("Resolve")
            resolve_btn.setFlat(True)
            resolve_btn.clicked.connect(lambda: self.resolve_requested.emit(self.comment))
            button_layout.addWidget(resolve_btn)

        edit_btn = QPushButton("Edit")
        edit_btn.setFlat(True)
        edit_btn.clicked.connect(lambda: self.edit_requested.emit(self.comment))
        button_layout.addWidget(edit_btn)

        delete_btn = QPushButton("Delete")
        delete_btn.setFlat(True)
        delete_btn.clicked.connect(lambda: self.delete_requested.emit(self.comment))
        button_layout.addWidget(delete_btn)

        layout.addLayout(button_layout)

        # Show if resolved
        if self.comment.get('resolved', False):
            self.setStyleSheet("QFrame { border: 1px solid #9c9; background-color: #f0f8f0; }")


class CommentsPanel(QWidget):
    """
    Comments panel for managing document comments and annotations.
    """

    # Signals
    comment_added = Signal(dict)  # Emits new comment data
    comment_edited = Signal(dict)  # Emits edited comment data
    comment_deleted = Signal(str)  # Emits comment ID
    comment_resolved = Signal(str)  # Emits comment ID
    comment_clicked = Signal(dict)  # Emits comment data for navigation

    def __init__(self, parent=None):
        super().__init__(parent)
        self.comments = []
        self.setup_ui()

    def setup_ui(self):
        """Initialize the panel UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)

        # Title and controls
        title_layout = QHBoxLayout()

        title_label = QLabel("Comments")
        title_font = QFont()
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_layout.addWidget(title_label)

        title_layout.addStretch()

        # New comment button
        self.new_btn = QPushButton("New")
        self.new_btn.clicked.connect(self.on_new_comment)
        title_layout.addWidget(self.new_btn)

        layout.addLayout(title_layout)

        # Filter controls
        filter_layout = QHBoxLayout()

        filter_layout.addWidget(QLabel("Show:"))

        self.filter_combo = QPushButton("All Comments")
        self.filter_combo.setFlat(True)
        filter_menu = QMenu()
        filter_menu.addAction("All Comments", lambda: self.set_filter("all"))
        filter_menu.addAction("Active", lambda: self.set_filter("active"))
        filter_menu.addAction("Resolved", lambda: self.set_filter("resolved"))
        self.filter_combo.setMenu(filter_menu)
        filter_layout.addWidget(self.filter_combo)

        filter_layout.addStretch()

        layout.addLayout(filter_layout)

        # Comments list
        self.comments_container = QWidget()
        self.comments_layout = QVBoxLayout(self.comments_container)
        self.comments_layout.setSpacing(10)
        self.comments_layout.addStretch()

        from PySide6.QtWidgets import QScrollArea
        scroll = QScrollArea()
        scroll.setWidget(self.comments_container)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        layout.addWidget(scroll)

        # Info label (shown when no comments)
        self.info_label = QLabel("No comments in document")
        self.info_label.setAlignment(Qt.AlignCenter)
        self.info_label.setStyleSheet("color: gray;")
        layout.addWidget(self.info_label)

        self.current_filter = "all"
        self.update_display()

    def add_comment(self, comment_data):
        """Add a new comment to the panel."""
        # Ensure comment has required fields
        if 'id' not in comment_data:
            comment_data['id'] = str(len(self.comments) + 1)
        if 'date' not in comment_data:
            comment_data['date'] = QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm")
        if 'author' not in comment_data:
            comment_data['author'] = "Current User"
        if 'resolved' not in comment_data:
            comment_data['resolved'] = False

        self.comments.append(comment_data)
        self.update_display()
        self.comment_added.emit(comment_data)

    def remove_comment(self, comment_id):
        """Remove a comment by ID."""
        self.comments = [c for c in self.comments if c.get('id') != comment_id]
        self.update_display()
        self.comment_deleted.emit(comment_id)

    def resolve_comment(self, comment_id):
        """Mark a comment as resolved."""
        for comment in self.comments:
            if comment.get('id') == comment_id:
                comment['resolved'] = True
                break
        self.update_display()
        self.comment_resolved.emit(comment_id)

    def set_filter(self, filter_type):
        """Set the comment filter."""
        self.current_filter = filter_type
        filter_labels = {
            "all": "All Comments",
            "active": "Active",
            "resolved": "Resolved"
        }
        self.filter_combo.setText(filter_labels.get(filter_type, "All Comments"))
        self.update_display()

    def update_display(self):
        """Update the comments display based on current filter."""
        # Clear existing widgets
        while self.comments_layout.count() > 1:  # Keep the stretch
            item = self.comments_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # Filter comments
        filtered_comments = self.comments
        if self.current_filter == "active":
            filtered_comments = [c for c in self.comments if not c.get('resolved', False)]
        elif self.current_filter == "resolved":
            filtered_comments = [c for c in self.comments if c.get('resolved', False)]

        # Show/hide info label
        if not filtered_comments:
            self.info_label.show()
            if self.current_filter == "active":
                self.info_label.setText("No active comments")
            elif self.current_filter == "resolved":
                self.info_label.setText("No resolved comments")
            else:
                self.info_label.setText("No comments in document")
        else:
            self.info_label.hide()

        # Add comment widgets
        for comment in filtered_comments:
            widget = CommentWidget(comment)
            widget.reply_requested.connect(self.on_reply_comment)
            widget.edit_requested.connect(self.on_edit_comment)
            widget.delete_requested.connect(self.on_delete_comment)
            widget.resolve_requested.connect(self.on_resolve_comment)
            self.comments_layout.insertWidget(self.comments_layout.count() - 1, widget)

    def on_new_comment(self):
        """Handle new comment button click."""
        # In a real implementation, this would show a dialog or inline editor
        sample_comment = {
            'text': 'New comment',
            'author': 'Current User',
            'context': 'Selected text or current paragraph',
        }
        self.add_comment(sample_comment)

    def on_reply_comment(self, comment):
        """Handle reply to comment."""
        # Create a reply (child comment)
        reply = {
            'text': 'Reply to comment',
            'author': 'Current User',
            'parent_id': comment.get('id'),
        }
        self.add_comment(reply)

    def on_edit_comment(self, comment):
        """Handle edit comment."""
        # In a real implementation, show edit dialog
        self.comment_edited.emit(comment)

    def on_delete_comment(self, comment):
        """Handle delete comment."""
        self.remove_comment(comment.get('id'))

    def on_resolve_comment(self, comment):
        """Handle resolve comment."""
        self.resolve_comment(comment.get('id'))

    def refresh(self):
        """Refresh the comments display."""
        self.update_display()
