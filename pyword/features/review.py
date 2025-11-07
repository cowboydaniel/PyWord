"""
Review Features for PyWord.

This module provides a unified review interface that integrates track changes,
comments, and document comparison for collaborative document editing.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                               QTabWidget, QWidget, QGroupBox, QTextEdit, QListWidget,
                               QListWidgetItem, QSplitter, QCheckBox, QComboBox,
                               QMessageBox, QMenu, QToolBar, QStatusBar, QFrame,
                               QScrollArea, QRadioButton, QButtonGroup)
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtGui import QAction, QIcon, QColor, QTextCharFormat, QFont
from datetime import datetime


class ReviewMode:
    """Review mode types."""
    EDIT = "Edit Mode"
    REVIEW = "Review Mode"
    FINAL = "Final Mode"


class ReviewManager:
    """Manages the review process for documents."""

    def __init__(self, parent):
        self.parent = parent
        self.mode = ReviewMode.EDIT
        self.track_changes_enabled = False
        self.show_comments = True
        self.show_markup = True
        self.reviewer_name = "Reviewer"

        # Integration with other features (to be set externally)
        self.track_changes_manager = None
        self.comments_manager = None

    def set_mode(self, mode):
        """Set the review mode."""
        self.mode = mode

        if mode == ReviewMode.EDIT:
            # Normal editing, all features available
            self.show_markup = True
            self.show_comments = True
        elif mode == ReviewMode.REVIEW:
            # Review mode, show all markup
            self.show_markup = True
            self.show_comments = True
        elif mode == ReviewMode.FINAL:
            # Final mode, hide all markup
            self.show_markup = False
            self.show_comments = False

        self._apply_mode_settings()

    def _apply_mode_settings(self):
        """Apply the current mode settings."""
        if self.track_changes_manager:
            self.track_changes_manager.show_changes = self.show_markup

        if self.comments_manager:
            self.comments_manager.show_comments = self.show_comments

    def enable_track_changes(self):
        """Enable track changes."""
        if self.track_changes_manager:
            self.track_changes_manager.enable_tracking(self.reviewer_name)
            self.track_changes_enabled = True

    def disable_track_changes(self):
        """Disable track changes."""
        if self.track_changes_manager:
            self.track_changes_manager.disable_tracking()
            self.track_changes_enabled = False

    def toggle_track_changes(self):
        """Toggle track changes."""
        if self.track_changes_enabled:
            self.disable_track_changes()
        else:
            self.enable_track_changes()

    def accept_all_changes(self):
        """Accept all pending changes."""
        if self.track_changes_manager:
            return self.track_changes_manager.accept_all_changes()
        return False

    def reject_all_changes(self):
        """Reject all pending changes."""
        if self.track_changes_manager:
            return self.track_changes_manager.reject_all_changes()
        return False

    def get_review_summary(self):
        """Get a summary of the review status."""
        summary = {
            'mode': self.mode,
            'track_changes_enabled': self.track_changes_enabled,
            'reviewer': self.reviewer_name,
            'pending_changes': 0,
            'active_comments': 0,
            'resolved_comments': 0
        }

        if self.track_changes_manager:
            summary['pending_changes'] = len(self.track_changes_manager.get_pending_changes())

        if self.comments_manager:
            summary['active_comments'] = len(self.comments_manager.get_active_comments())
            summary['resolved_comments'] = len(self.comments_manager.get_resolved_comments())

        return summary


class ReviewPanel(QWidget):
    """Panel for review operations."""

    mode_changed = Signal(str)

    def __init__(self, review_manager, parent=None):
        super().__init__(parent)
        self.review_manager = review_manager
        self.setup_ui()

    def setup_ui(self):
        """Setup the panel UI."""
        layout = QVBoxLayout()

        # Header
        header = QLabel("<b>Review</b>")
        header_font = QFont()
        header_font.setPointSize(12)
        header.setFont(header_font)
        layout.addWidget(header)

        # Mode selection
        mode_group = QGroupBox("Review Mode")
        mode_layout = QVBoxLayout()

        self.mode_button_group = QButtonGroup()

        self.edit_mode_radio = QRadioButton("Edit Mode")
        self.edit_mode_radio.setToolTip("Normal editing with markup visible")
        self.review_mode_radio = QRadioButton("Review Mode")
        self.review_mode_radio.setToolTip("Review changes and comments")
        self.final_mode_radio = QRadioButton("Final Mode")
        self.final_mode_radio.setToolTip("Hide all markup")

        self.mode_button_group.addButton(self.edit_mode_radio)
        self.mode_button_group.addButton(self.review_mode_radio)
        self.mode_button_group.addButton(self.final_mode_radio)

        self.edit_mode_radio.setChecked(True)
        self.edit_mode_radio.toggled.connect(self.on_mode_changed)
        self.review_mode_radio.toggled.connect(self.on_mode_changed)
        self.final_mode_radio.toggled.connect(self.on_mode_changed)

        mode_layout.addWidget(self.edit_mode_radio)
        mode_layout.addWidget(self.review_mode_radio)
        mode_layout.addWidget(self.final_mode_radio)

        mode_group.setLayout(mode_layout)
        layout.addWidget(mode_group)

        # Track changes controls
        tracking_group = QGroupBox("Track Changes")
        tracking_layout = QVBoxLayout()

        self.track_changes_button = QPushButton("Enable Track Changes")
        self.track_changes_button.clicked.connect(self.toggle_track_changes)
        tracking_layout.addWidget(self.track_changes_button)

        self.show_changes_checkbox = QCheckBox("Show Changes")
        self.show_changes_checkbox.setChecked(True)
        self.show_changes_checkbox.toggled.connect(self.toggle_show_changes)
        tracking_layout.addWidget(self.show_changes_checkbox)

        tracking_group.setLayout(tracking_layout)
        layout.addWidget(tracking_group)

        # Quick actions
        actions_group = QGroupBox("Quick Actions")
        actions_layout = QVBoxLayout()

        accept_all_button = QPushButton("Accept All Changes")
        accept_all_button.clicked.connect(self.accept_all_changes)
        actions_layout.addWidget(accept_all_button)

        reject_all_button = QPushButton("Reject All Changes")
        reject_all_button.clicked.connect(self.reject_all_changes)
        actions_layout.addWidget(reject_all_button)

        actions_group.setLayout(actions_layout)
        layout.addWidget(actions_group)

        # Summary
        summary_group = QGroupBox("Summary")
        summary_layout = QVBoxLayout()

        self.summary_label = QLabel()
        self.update_summary()
        summary_layout.addWidget(self.summary_label)

        refresh_button = QPushButton("Refresh")
        refresh_button.clicked.connect(self.update_summary)
        summary_layout.addWidget(refresh_button)

        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)

        layout.addStretch()
        self.setLayout(layout)

    def on_mode_changed(self):
        """Handle mode change."""
        if self.edit_mode_radio.isChecked():
            self.review_manager.set_mode(ReviewMode.EDIT)
            self.mode_changed.emit(ReviewMode.EDIT)
        elif self.review_mode_radio.isChecked():
            self.review_manager.set_mode(ReviewMode.REVIEW)
            self.mode_changed.emit(ReviewMode.REVIEW)
        elif self.final_mode_radio.isChecked():
            self.review_manager.set_mode(ReviewMode.FINAL)
            self.mode_changed.emit(ReviewMode.FINAL)

    def toggle_track_changes(self):
        """Toggle track changes."""
        self.review_manager.toggle_track_changes()

        if self.review_manager.track_changes_enabled:
            self.track_changes_button.setText("Disable Track Changes")
        else:
            self.track_changes_button.setText("Enable Track Changes")

        self.update_summary()

    def toggle_show_changes(self):
        """Toggle showing changes."""
        if self.review_manager.track_changes_manager:
            self.review_manager.track_changes_manager.toggle_show_changes()

    def accept_all_changes(self):
        """Accept all changes."""
        reply = QMessageBox.question(
            self,
            "Accept All Changes",
            "Are you sure you want to accept all pending changes?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.review_manager.accept_all_changes()
            self.update_summary()
            QMessageBox.information(self, "Success", "All changes accepted!")

    def reject_all_changes(self):
        """Reject all changes."""
        reply = QMessageBox.question(
            self,
            "Reject All Changes",
            "Are you sure you want to reject all pending changes?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.review_manager.reject_all_changes()
            self.update_summary()
            QMessageBox.information(self, "Success", "All changes rejected!")

    def update_summary(self):
        """Update the summary display."""
        summary = self.review_manager.get_review_summary()

        summary_text = f"<b>Current Mode:</b> {summary['mode']}<br>"
        summary_text += f"<b>Reviewer:</b> {summary['reviewer']}<br>"
        summary_text += f"<b>Track Changes:</b> {'Enabled' if summary['track_changes_enabled'] else 'Disabled'}<br>"
        summary_text += f"<b>Pending Changes:</b> {summary['pending_changes']}<br>"
        summary_text += f"<b>Active Comments:</b> {summary['active_comments']}<br>"
        summary_text += f"<b>Resolved Comments:</b> {summary['resolved_comments']}"

        self.summary_label.setText(summary_text)


class ReviewDialog(QDialog):
    """Main dialog for document review."""

    def __init__(self, review_manager, parent=None):
        super().__init__(parent)
        self.review_manager = review_manager
        self.setWindowTitle("Review Document")
        self.setMinimumSize(800, 600)

        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Toolbar
        toolbar = QToolBar()

        # Track changes actions
        self.track_changes_action = QAction("Track Changes", self)
        self.track_changes_action.setCheckable(True)
        self.track_changes_action.toggled.connect(self.toggle_track_changes)
        toolbar.addAction(self.track_changes_action)

        toolbar.addSeparator()

        # Show markup
        self.show_markup_action = QAction("Show Markup", self)
        self.show_markup_action.setCheckable(True)
        self.show_markup_action.setChecked(True)
        self.show_markup_action.toggled.connect(self.toggle_show_markup)
        toolbar.addAction(self.show_markup_action)

        # Show comments
        self.show_comments_action = QAction("Show Comments", self)
        self.show_comments_action.setCheckable(True)
        self.show_comments_action.setChecked(True)
        self.show_comments_action.toggled.connect(self.toggle_show_comments)
        toolbar.addAction(self.show_comments_action)

        toolbar.addSeparator()

        # Accept/Reject
        accept_action = QAction("Accept Change", self)
        accept_action.triggered.connect(self.accept_current_change)
        toolbar.addAction(accept_action)

        reject_action = QAction("Reject Change", self)
        reject_action.triggered.connect(self.reject_current_change)
        toolbar.addAction(reject_action)

        layout.addWidget(toolbar)

        # Main content
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Review panel
        self.review_panel = ReviewPanel(self.review_manager)
        splitter.addWidget(self.review_panel)

        # Tabs for details
        tabs = QTabWidget()

        # Changes tab
        changes_tab = self.create_changes_tab()
        tabs.addTab(changes_tab, "Changes")

        # Comments tab
        comments_tab = self.create_comments_tab()
        tabs.addTab(comments_tab, "Comments")

        # Statistics tab
        stats_tab = self.create_statistics_tab()
        tabs.addTab(stats_tab, "Statistics")

        splitter.addWidget(tabs)
        splitter.setSizes([250, 550])

        layout.addWidget(splitter)

        # Status bar
        status_bar = QStatusBar()
        self.status_label = QLabel("Ready")
        status_bar.addWidget(self.status_label)
        layout.addWidget(status_bar)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def create_changes_tab(self):
        """Create the changes tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Filter
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Show:"))

        self.changes_filter = QComboBox()
        self.changes_filter.addItems(["All Changes", "Insertions", "Deletions", "Formatting"])
        self.changes_filter.currentTextChanged.connect(self.refresh_changes)
        filter_layout.addWidget(self.changes_filter)

        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Changes list
        self.changes_list = QListWidget()
        self.changes_list.itemSelectionChanged.connect(self.on_change_selected)
        layout.addWidget(self.changes_list)

        # Change details
        self.change_details = QTextEdit()
        self.change_details.setReadOnly(True)
        self.change_details.setMaximumHeight(150)
        layout.addWidget(self.change_details)

        # Action buttons
        action_layout = QHBoxLayout()

        accept_button = QPushButton("Accept")
        accept_button.clicked.connect(self.accept_selected_change)
        action_layout.addWidget(accept_button)

        reject_button = QPushButton("Reject")
        reject_button.clicked.connect(self.reject_selected_change)
        action_layout.addWidget(reject_button)

        action_layout.addStretch()
        layout.addLayout(action_layout)

        tab.setLayout(layout)
        return tab

    def create_comments_tab(self):
        """Create the comments tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Filter
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Show:"))

        self.comments_filter = QComboBox()
        self.comments_filter.addItems(["All Comments", "Active", "Resolved"])
        self.comments_filter.currentTextChanged.connect(self.refresh_comments)
        filter_layout.addWidget(self.comments_filter)

        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Comments list
        self.comments_list = QListWidget()
        layout.addWidget(self.comments_list)

        tab.setLayout(layout)
        return tab

    def create_statistics_tab(self):
        """Create the statistics tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        self.statistics_text = QTextEdit()
        self.statistics_text.setReadOnly(True)
        layout.addWidget(self.statistics_text)

        refresh_button = QPushButton("Refresh Statistics")
        refresh_button.clicked.connect(self.update_statistics)
        layout.addWidget(refresh_button)

        tab.setLayout(layout)
        return tab

    def toggle_track_changes(self, checked):
        """Toggle track changes."""
        if checked:
            self.review_manager.enable_track_changes()
        else:
            self.review_manager.disable_track_changes()

        self.update_status()

    def toggle_show_markup(self, checked):
        """Toggle showing markup."""
        self.review_manager.show_markup = checked
        self.review_manager._apply_mode_settings()

    def toggle_show_comments(self, checked):
        """Toggle showing comments."""
        self.review_manager.show_comments = checked
        self.review_manager._apply_mode_settings()

    def accept_current_change(self):
        """Accept the current/selected change."""
        self.accept_selected_change()

    def reject_current_change(self):
        """Reject the current/selected change."""
        self.reject_selected_change()

    def refresh_changes(self):
        """Refresh the changes list."""
        self.changes_list.clear()

        if not self.review_manager.track_changes_manager:
            return

        changes = self.review_manager.track_changes_manager.get_pending_changes()

        # Apply filter
        filter_text = self.changes_filter.currentText()
        if filter_text == "Insertions":
            changes = [c for c in changes if c.change_type == "Insertion"]
        elif filter_text == "Deletions":
            changes = [c for c in changes if c.change_type == "Deletion"]
        elif filter_text == "Formatting":
            changes = [c for c in changes if c.change_type == "Formatting"]

        for change in changes:
            item_text = f"{change.change_type} by {change.author}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, change)
            self.changes_list.addItem(item)

    def refresh_comments(self):
        """Refresh the comments list."""
        self.comments_list.clear()

        if not self.review_manager.comments_manager:
            return

        filter_text = self.comments_filter.currentText()

        if filter_text == "All Comments":
            comments = self.review_manager.comments_manager.get_all_root_comments()
        elif filter_text == "Active":
            comments = self.review_manager.comments_manager.get_active_comments()
        elif filter_text == "Resolved":
            comments = self.review_manager.comments_manager.get_resolved_comments()
        else:
            comments = []

        for comment in comments:
            preview = comment.text[:50]
            if len(comment.text) > 50:
                preview += "..."
            item_text = f"{comment.author}: {preview}"
            item = QListWidgetItem(item_text)
            self.comments_list.addItem(item)

    def on_change_selected(self):
        """Handle change selection."""
        items = self.changes_list.selectedItems()
        if not items:
            return

        change = items[0].data(Qt.ItemDataRole.UserRole)

        details = f"<b>Type:</b> {change.change_type}<br>"
        details += f"<b>Author:</b> {change.author}<br>"
        details += f"<b>Time:</b> {change.timestamp.strftime('%Y-%m-%d %H:%M:%S')}<br>"
        details += f"<b>Content:</b> {change.content}"

        self.change_details.setHtml(details)

    def accept_selected_change(self):
        """Accept the selected change."""
        items = self.changes_list.selectedItems()
        if not items:
            return

        change = items[0].data(Qt.ItemDataRole.UserRole)
        if self.review_manager.track_changes_manager:
            self.review_manager.track_changes_manager.accept_change(change)
            self.refresh_changes()
            self.update_status()

    def reject_selected_change(self):
        """Reject the selected change."""
        items = self.changes_list.selectedItems()
        if not items:
            return

        change = items[0].data(Qt.ItemDataRole.UserRole)
        if self.review_manager.track_changes_manager:
            self.review_manager.track_changes_manager.reject_change(change)
            self.refresh_changes()
            self.update_status()

    def update_statistics(self):
        """Update the statistics display."""
        stats_html = "<h3>Review Statistics</h3>"

        summary = self.review_manager.get_review_summary()

        stats_html += f"<p><b>Review Mode:</b> {summary['mode']}</p>"
        stats_html += f"<p><b>Reviewer:</b> {summary['reviewer']}</p>"
        stats_html += f"<p><b>Track Changes:</b> {'Enabled' if summary['track_changes_enabled'] else 'Disabled'}</p>"
        stats_html += "<hr>"
        stats_html += f"<p><b>Pending Changes:</b> {summary['pending_changes']}</p>"
        stats_html += f"<p><b>Active Comments:</b> {summary['active_comments']}</p>"
        stats_html += f"<p><b>Resolved Comments:</b> {summary['resolved_comments']}</p>"
        stats_html += "<hr>"
        stats_html += f"<p><b>Last Updated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>"

        self.statistics_text.setHtml(stats_html)

    def update_status(self):
        """Update the status bar."""
        summary = self.review_manager.get_review_summary()
        status = f"Mode: {summary['mode']} | Changes: {summary['pending_changes']} | Comments: {summary['active_comments']}"
        self.status_label.setText(status)
