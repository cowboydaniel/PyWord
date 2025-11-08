"""Quick Access Toolbar - customizable toolbar for frequently used commands."""

from PySide6.QtWidgets import (QToolBar, QToolButton, QMenu, QDialog, QVBoxLayout,
                               QHBoxLayout, QListWidget, QListWidgetItem, QPushButton,
                               QLabel, QSplitter)
from PySide6.QtCore import Qt, Signal, QSettings, QSize
from PySide6.QtGui import QIcon, QAction


class QuickAccessToolbar(QToolBar):
    """Customizable quick access toolbar for frequently used commands."""

    def __init__(self, parent=None):
        super().__init__("Quick Access Toolbar", parent)
        self.setObjectName("QuickAccessToolbar")
        self.setMovable(False)
        self.setFloatable(False)
        self.setIconSize(QSize(16, 16))
        self.settings = QSettings("PyWord", "QuickAccessToolbar")

        # Default commands
        self.available_commands = {
            'save': {'name': 'Save', 'icon': QIcon(), 'tooltip': 'Save document'},
            'undo': {'name': 'Undo', 'icon': QIcon(), 'tooltip': 'Undo last action'},
            'redo': {'name': 'Redo', 'icon': QIcon(), 'tooltip': 'Redo last action'},
            'print': {'name': 'Print', 'icon': QIcon(), 'tooltip': 'Print document'},
            'print_preview': {'name': 'Print Preview', 'icon': QIcon(), 'tooltip': 'Preview before printing'},
            'new': {'name': 'New', 'icon': QIcon(), 'tooltip': 'New document'},
            'open': {'name': 'Open', 'icon': QIcon(), 'tooltip': 'Open document'},
            'save_as': {'name': 'Save As', 'icon': QIcon(), 'tooltip': 'Save document as'},
            'cut': {'name': 'Cut', 'icon': QIcon(), 'tooltip': 'Cut to clipboard'},
            'copy': {'name': 'Copy', 'icon': QIcon(), 'tooltip': 'Copy to clipboard'},
            'paste': {'name': 'Paste', 'icon': QIcon(), 'tooltip': 'Paste from clipboard'},
            'bold': {'name': 'Bold', 'icon': QIcon(), 'tooltip': 'Bold text'},
            'italic': {'name': 'Italic', 'icon': QIcon(), 'tooltip': 'Italic text'},
            'underline': {'name': 'Underline', 'icon': QIcon(), 'tooltip': 'Underline text'},
        }

        self.command_actions = {}
        self.setup_ui()
        self.load_configuration()

    def setup_ui(self):
        """Initialize the toolbar UI."""
        self.setStyleSheet("""
            QToolBar {
                background: #f3f3f3;
                border: none;
                spacing: 2px;
                padding: 2px;
            }
            QToolButton {
                border: 1px solid transparent;
                border-radius: 3px;
                padding: 3px;
                background: transparent;
            }
            QToolButton:hover {
                background: rgba(0, 120, 215, 0.1);
                border: 1px solid rgba(0, 120, 215, 0.3);
            }
            QToolButton:pressed {
                background: rgba(0, 120, 215, 0.2);
            }
        """)

        # Add customize button
        self.customize_button = QToolButton()
        self.customize_button.setText("...")
        self.customize_button.setToolTip("Customize Quick Access Toolbar")
        self.customize_button.clicked.connect(self.show_customize_dialog)
        self.addWidget(self.customize_button)

        self.addSeparator()

    def add_command(self, command_id: str, action: QAction = None):
        """Add a command to the toolbar."""
        if command_id not in self.available_commands:
            return

        cmd = self.available_commands[command_id]

        if action:
            # Use existing action
            button = QToolButton()
            button.setDefaultAction(action)
            button.setToolButtonStyle(Qt.ToolButtonIconOnly)
            self.addWidget(button)
            self.command_actions[command_id] = action
        else:
            # Create new action
            new_action = QAction(cmd['icon'], cmd['name'], self)
            new_action.setToolTip(cmd['tooltip'])
            button = QToolButton()
            button.setDefaultAction(new_action)
            button.setToolButtonStyle(Qt.ToolButtonIconOnly)
            self.addWidget(button)
            self.command_actions[command_id] = new_action

    def remove_command(self, command_id: str):
        """Remove a command from the toolbar."""
        if command_id in self.command_actions:
            action = self.command_actions[command_id]
            self.removeAction(action)
            del self.command_actions[command_id]

    def get_active_commands(self) -> list:
        """Get list of currently active command IDs."""
        return list(self.command_actions.keys())

    def set_commands(self, command_ids: list, actions: dict = None):
        """Set the toolbar commands."""
        # Clear existing commands
        for cmd_id in list(self.command_actions.keys()):
            self.remove_command(cmd_id)

        # Add new commands
        for cmd_id in command_ids:
            action = actions.get(cmd_id) if actions else None
            self.add_command(cmd_id, action)

    def save_configuration(self):
        """Save toolbar configuration to settings."""
        active_commands = self.get_active_commands()
        self.settings.setValue("commands", active_commands)

    def load_configuration(self):
        """Load toolbar configuration from settings."""
        # Default commands if none saved
        default_commands = ['save', 'undo', 'redo', 'print']
        saved_commands = self.settings.value("commands", default_commands)

        if saved_commands:
            self.set_commands(saved_commands)

    def show_customize_dialog(self):
        """Show the customization dialog."""
        dialog = QuickAccessCustomizeDialog(self, self.parent())
        if dialog.exec() == QDialog.Accepted:
            selected_commands = dialog.get_selected_commands()
            self.set_commands(selected_commands)
            self.save_configuration()

    def connect_command(self, command_id: str, callback):
        """Connect a command to a callback function."""
        if command_id in self.command_actions:
            self.command_actions[command_id].triggered.connect(callback)


class QuickAccessCustomizeDialog(QDialog):
    """Dialog for customizing the Quick Access Toolbar."""

    def __init__(self, toolbar: QuickAccessToolbar, parent=None):
        super().__init__(parent)
        self.toolbar = toolbar
        self.setWindowTitle("Customize Quick Access Toolbar")
        self.setModal(True)
        self.resize(600, 400)
        self.setup_ui()

    def setup_ui(self):
        """Initialize the dialog UI."""
        layout = QVBoxLayout(self)

        # Instructions
        instructions = QLabel("Choose commands to add to the Quick Access Toolbar:")
        layout.addWidget(instructions)

        # Splitter for two lists
        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter)

        # Available commands list
        available_container = QWidget()
        available_layout = QVBoxLayout(available_container)
        available_layout.addWidget(QLabel("Available Commands:"))

        self.available_list = QListWidget()
        self.available_list.setSelectionMode(QListWidget.MultiSelection)
        available_layout.addWidget(self.available_list)

        splitter.addWidget(available_container)

        # Buttons for moving items
        button_container = QWidget()
        button_layout = QVBoxLayout(button_container)
        button_layout.addStretch()

        self.add_button = QPushButton("Add >>")
        self.add_button.clicked.connect(self.add_commands)
        button_layout.addWidget(self.add_button)

        self.remove_button = QPushButton("<< Remove")
        self.remove_button.clicked.connect(self.remove_commands)
        button_layout.addWidget(self.remove_button)

        button_layout.addStretch()

        self.move_up_button = QPushButton("Move Up")
        self.move_up_button.clicked.connect(self.move_up)
        button_layout.addWidget(self.move_up_button)

        self.move_down_button = QPushButton("Move Down")
        self.move_down_button.clicked.connect(self.move_down)
        button_layout.addWidget(self.move_down_button)

        button_layout.addStretch()

        splitter.addWidget(button_container)

        # Active commands list
        active_container = QWidget()
        active_layout = QVBoxLayout(active_container)
        active_layout.addWidget(QLabel("Quick Access Toolbar:"))

        self.active_list = QListWidget()
        active_layout.addWidget(self.active_list)

        splitter.addWidget(active_container)

        # Dialog buttons
        button_box = QHBoxLayout()
        button_box.addStretch()

        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        button_box.addWidget(ok_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_box.addWidget(cancel_button)

        layout.addLayout(button_box)

        # Populate lists
        self.populate_lists()

    def populate_lists(self):
        """Populate the available and active command lists."""
        # Get current commands
        active_commands = self.toolbar.get_active_commands()

        # Populate available commands
        for cmd_id, cmd_data in self.toolbar.available_commands.items():
            if cmd_id not in active_commands:
                item = QListWidgetItem(cmd_data['name'])
                item.setData(Qt.UserRole, cmd_id)
                self.available_list.addItem(item)

        # Populate active commands
        for cmd_id in active_commands:
            if cmd_id in self.toolbar.available_commands:
                cmd_data = self.toolbar.available_commands[cmd_id]
                item = QListWidgetItem(cmd_data['name'])
                item.setData(Qt.UserRole, cmd_id)
                self.active_list.addItem(item)

    def add_commands(self):
        """Move selected commands from available to active."""
        selected_items = self.available_list.selectedItems()
        for item in selected_items:
            cmd_id = item.data(Qt.UserRole)
            cmd_name = item.text()

            # Add to active list
            new_item = QListWidgetItem(cmd_name)
            new_item.setData(Qt.UserRole, cmd_id)
            self.active_list.addItem(new_item)

            # Remove from available list
            row = self.available_list.row(item)
            self.available_list.takeItem(row)

    def remove_commands(self):
        """Move selected commands from active to available."""
        selected_items = self.active_list.selectedItems()
        for item in selected_items:
            cmd_id = item.data(Qt.UserRole)
            cmd_name = item.text()

            # Add to available list
            new_item = QListWidgetItem(cmd_name)
            new_item.setData(Qt.UserRole, cmd_id)
            self.available_list.addItem(new_item)

            # Remove from active list
            row = self.active_list.row(item)
            self.active_list.takeItem(row)

    def move_up(self):
        """Move selected command up in the active list."""
        current_row = self.active_list.currentRow()
        if current_row > 0:
            item = self.active_list.takeItem(current_row)
            self.active_list.insertItem(current_row - 1, item)
            self.active_list.setCurrentRow(current_row - 1)

    def move_down(self):
        """Move selected command down in the active list."""
        current_row = self.active_list.currentRow()
        if current_row < self.active_list.count() - 1:
            item = self.active_list.takeItem(current_row)
            self.active_list.insertItem(current_row + 1, item)
            self.active_list.setCurrentRow(current_row + 1)

    def get_selected_commands(self) -> list:
        """Get the list of selected command IDs in order."""
        commands = []
        for i in range(self.active_list.count()):
            item = self.active_list.item(i)
            cmd_id = item.data(Qt.UserRole)
            commands.append(cmd_id)
        return commands
