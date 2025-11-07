"""Tabs Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QListWidget, QListWidgetItem,
                             QGroupBox, QFormLayout, QSpinBox, QComboBox,
                             QRadioButton, QButtonGroup)
from PySide6.QtCore import Qt
from .base_dialog import BaseDialog


class TabsDialog(BaseDialog):
    """Dialog for configuring tab stops."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Tabs")
        self.tab_stops = []
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Tab stop position
        position_group = QGroupBox("Tab Stop Position")
        position_layout = QHBoxLayout()

        self.position_spinbox = QSpinBox()
        self.position_spinbox.setMinimum(0)
        self.position_spinbox.setMaximum(1000)
        self.position_spinbox.setValue(0)
        self.position_spinbox.setSuffix(" pt")
        position_layout.addWidget(self.position_spinbox)

        position_group.setLayout(position_layout)
        layout.addWidget(position_group)

        # Alignment
        alignment_group = QGroupBox("Alignment")
        alignment_layout = QVBoxLayout()

        self.alignment_group = QButtonGroup()
        alignments = ["Left", "Center", "Right", "Decimal", "Bar"]

        for i, alignment in enumerate(alignments):
            radio = QRadioButton(alignment)
            if i == 0:
                radio.setChecked(True)
            self.alignment_group.addButton(radio, i)
            alignment_layout.addWidget(radio)

        alignment_group.setLayout(alignment_layout)
        layout.addWidget(alignment_group)

        # Leader
        leader_group = QGroupBox("Leader")
        leader_layout = QVBoxLayout()

        self.leader_group = QButtonGroup()
        leaders = ["None", "....", "----", "____"]

        for i, leader in enumerate(leaders):
            radio = QRadioButton(leader)
            if i == 0:
                radio.setChecked(True)
            self.leader_group.addButton(radio, i)
            leader_layout.addWidget(radio)

        leader_group.setLayout(leader_layout)
        layout.addWidget(leader_group)

        # Tab stops list
        list_group = QGroupBox("Tab Stops")
        list_layout = QVBoxLayout()

        self.tabs_list = QListWidget()
        list_layout.addWidget(self.tabs_list)

        # List buttons
        list_buttons = QHBoxLayout()

        self.set_button = QPushButton("Set")
        self.set_button.clicked.connect(self.set_tab_stop)
        list_buttons.addWidget(self.set_button)

        self.clear_button = QPushButton("Clear")
        self.clear_button.clicked.connect(self.clear_tab_stop)
        list_buttons.addWidget(self.clear_button)

        self.clear_all_button = QPushButton("Clear All")
        self.clear_all_button.clicked.connect(self.clear_all_tab_stops)
        list_buttons.addWidget(self.clear_all_button)

        list_layout.addLayout(list_buttons)
        list_group.setLayout(list_layout)
        layout.addWidget(list_group)

        # Default tab stops
        default_group = QGroupBox("Default Tab Stops")
        default_layout = QFormLayout()

        self.default_spinbox = QSpinBox()
        self.default_spinbox.setMinimum(10)
        self.default_spinbox.setMaximum(1000)
        self.default_spinbox.setValue(36)  # Default 0.5 inches (36 pt)
        self.default_spinbox.setSuffix(" pt")
        default_layout.addRow("Default tab stops:", self.default_spinbox)

        default_group.setLayout(default_layout)
        layout.addWidget(default_group)

        # Dialog buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)
        button_layout.addWidget(self.ok_button)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.resize(350, 550)

    def set_tab_stop(self):
        """Add or update a tab stop."""
        position = self.position_spinbox.value()
        alignment = self.get_alignment()
        leader = self.get_leader()

        # Check if tab stop already exists at this position
        for i in range(self.tabs_list.count()):
            item = self.tabs_list.item(i)
            if item.data(Qt.UserRole) == position:
                # Update existing tab stop
                item.setText(f"{position} pt - {alignment} - {leader}")
                return

        # Add new tab stop
        item = QListWidgetItem(f"{position} pt - {alignment} - {leader}")
        item.setData(Qt.UserRole, position)
        self.tabs_list.addItem(item)
        self.tab_stops.append((position, alignment, leader))

    def clear_tab_stop(self):
        """Remove the selected tab stop."""
        current_item = self.tabs_list.currentItem()
        if current_item:
            position = current_item.data(Qt.UserRole)
            self.tab_stops = [ts for ts in self.tab_stops if ts[0] != position]
            self.tabs_list.takeItem(self.tabs_list.row(current_item))

    def clear_all_tab_stops(self):
        """Remove all tab stops."""
        self.tabs_list.clear()
        self.tab_stops.clear()

    def get_alignment(self):
        """Get the selected alignment."""
        button_id = self.alignment_group.checkedId()
        alignments = ["Left", "Center", "Right", "Decimal", "Bar"]
        return alignments[button_id] if 0 <= button_id < len(alignments) else "Left"

    def get_leader(self):
        """Get the selected leader."""
        button_id = self.leader_group.checkedId()
        leaders = ["None", "....", "----", "____"]
        return leaders[button_id] if 0 <= button_id < len(leaders) else "None"

    def get_tab_stops(self):
        """Get all configured tab stops."""
        return self.tab_stops

    def get_default_tab_stop(self):
        """Get the default tab stop spacing."""
        return self.default_spinbox.value()
