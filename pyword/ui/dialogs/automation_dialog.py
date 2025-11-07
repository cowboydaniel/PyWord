"""
Automation dialogs for PyWord.

Provides UI dialogs for automation features including:
- Macros (recording and management)
- Add-ins (plugins)
- Custom XML data storage
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QTextEdit, QListWidget, QListWidgetItem,
                               QMessageBox, QComboBox, QCheckBox, QGroupBox,
                               QTableWidget, QTableWidgetItem, QHeaderView,
                               QTabWidget, QWidget, QSplitter, QFileDialog,
                               QTreeWidget, QTreeWidgetItem, QFormLayout)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QColor, QKeySequence
from ...features.automation import (AutomationManager, Macro, MacroType, ActionType,
                                     AddIn, CustomXMLPart)


class MacroManagerDialog(QDialog):
    """Dialog for managing macros."""

    def __init__(self, automation_manager: AutomationManager, parent=None):
        super().__init__(parent)
        self.automation_manager = automation_manager
        self.setWindowTitle("Macro Manager")
        self.setModal(True)
        self.setMinimumSize(700, 500)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Macro list
        list_layout = QHBoxLayout()

        self.macro_list = QListWidget()
        self.macro_list.setMinimumWidth(200)
        self.macro_list.currentItemChanged.connect(self.on_macro_selected)
        self.refresh_macro_list()

        list_layout.addWidget(self.macro_list)

        # Macro details
        details_widget = QWidget()
        details_layout = QVBoxLayout()

        self.name_label = QLabel("Name: -")
        self.type_label = QLabel("Type: -")
        self.created_label = QLabel("Created: -")
        self.run_count_label = QLabel("Run Count: 0")
        self.shortcut_label = QLabel("Keyboard Shortcut: None")

        self.description_edit = QTextEdit()
        self.description_edit.setPlaceholderText("Macro description...")
        self.description_edit.setMaximumHeight(80)

        details_layout.addWidget(self.name_label)
        details_layout.addWidget(self.type_label)
        details_layout.addWidget(self.created_label)
        details_layout.addWidget(self.run_count_label)
        details_layout.addWidget(self.shortcut_label)
        details_layout.addWidget(QLabel("Description:"))
        details_layout.addWidget(self.description_edit)

        # Auto-run options
        auto_run_group = QGroupBox("Auto-Run Options")
        auto_run_layout = QVBoxLayout()

        self.auto_run_open_check = QCheckBox("Run on document open")
        self.auto_run_save_check = QCheckBox("Run on document save")

        auto_run_layout.addWidget(self.auto_run_open_check)
        auto_run_layout.addWidget(self.auto_run_save_check)

        auto_run_group.setLayout(auto_run_layout)
        details_layout.addWidget(auto_run_group)

        details_layout.addStretch()

        details_widget.setLayout(details_layout)
        list_layout.addWidget(details_widget, 1)

        layout.addLayout(list_layout)

        # Buttons
        button_layout = QHBoxLayout()

        self.new_button = QPushButton("New")
        self.new_button.clicked.connect(self.new_macro)

        self.record_button = QPushButton("Record")
        self.record_button.clicked.connect(self.record_macro)

        self.edit_button = QPushButton("Edit")
        self.edit_button.clicked.connect(self.edit_macro)

        self.run_button = QPushButton("Run")
        self.run_button.clicked.connect(self.run_macro)

        self.delete_button = QPushButton("Delete")
        self.delete_button.clicked.connect(self.delete_macro)

        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.accept)

        button_layout.addWidget(self.new_button)
        button_layout.addWidget(self.record_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.run_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addStretch()
        button_layout.addWidget(self.close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def refresh_macro_list(self):
        self.macro_list.clear()
        macros = self.automation_manager.get_all_macros()

        for macro in macros:
            item = QListWidgetItem(macro.name)
            item.setData(Qt.ItemDataRole.UserRole, macro.id)
            if not macro.enabled:
                item.setForeground(QColor("gray"))
            self.macro_list.addItem(item)

    def on_macro_selected(self, current, previous):
        if not current:
            return

        macro_id = current.data(Qt.ItemDataRole.UserRole)
        macro = self.automation_manager.get_macro(macro_id)

        if macro:
            self.name_label.setText(f"Name: {macro.name}")
            self.type_label.setText(f"Type: {macro.macro_type.value}")
            self.created_label.setText(f"Created: {macro.created_date.strftime('%Y-%m-%d %H:%M')}")
            self.run_count_label.setText(f"Run Count: {macro.run_count}")
            self.shortcut_label.setText(f"Keyboard Shortcut: {macro.keyboard_shortcut or 'None'}")
            self.description_edit.setPlainText(macro.description)
            self.auto_run_open_check.setChecked(macro.auto_run_on_open)
            self.auto_run_save_check.setChecked(macro.auto_run_on_save)

    def new_macro(self):
        dialog = EditMacroDialog(self.automation_manager, None, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_macro_list()

    def record_macro(self):
        dialog = RecordMacroDialog(self.automation_manager, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_macro_list()

    def edit_macro(self):
        item = self.macro_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Error", "Please select a macro to edit!")
            return

        macro_id = item.data(Qt.ItemDataRole.UserRole)
        macro = self.automation_manager.get_macro(macro_id)

        dialog = EditMacroDialog(self.automation_manager, macro, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_macro_list()

    def run_macro(self):
        item = self.macro_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Error", "Please select a macro to run!")
            return

        macro_id = item.data(Qt.ItemDataRole.UserRole)

        if self.automation_manager.run_macro(macro_id):
            QMessageBox.information(self, "Success", "Macro executed successfully!")
            self.refresh_macro_list()
        else:
            QMessageBox.critical(self, "Error", "Failed to execute macro!")

    def delete_macro(self):
        item = self.macro_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Error", "Please select a macro to delete!")
            return

        reply = QMessageBox.question(self, "Confirm",
            "Are you sure you want to delete this macro?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            macro_id = item.data(Qt.ItemDataRole.UserRole)
            self.automation_manager.delete_macro(macro_id)
            self.refresh_macro_list()


class RecordMacroDialog(QDialog):
    """Dialog for recording a macro."""

    def __init__(self, automation_manager: AutomationManager, parent=None):
        super().__init__(parent)
        self.automation_manager = automation_manager
        self.setWindowTitle("Record Macro")
        self.setModal(True)
        self.setMinimumWidth(400)
        self.recording_macro = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Macro name
        form_layout = QFormLayout()

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Enter macro name")

        self.description_edit = QLineEdit()
        self.description_edit.setPlaceholderText("Optional description")

        form_layout.addRow("Macro Name:", self.name_edit)
        form_layout.addRow("Description:", self.description_edit)

        layout.addLayout(form_layout)

        # Recording status
        self.status_label = QLabel("Ready to record")
        self.status_label.setStyleSheet("QLabel { color: blue; font-weight: bold; }")
        layout.addWidget(self.status_label)

        # Info
        info_label = QLabel("\nClick 'Start Recording' to begin.\n"
                          "Perform the actions you want to record.\n"
                          "Click 'Stop Recording' when done.")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        # Buttons
        button_layout = QHBoxLayout()

        self.start_button = QPushButton("Start Recording")
        self.start_button.clicked.connect(self.start_recording)

        self.stop_button = QPushButton("Stop Recording")
        self.stop_button.clicked.connect(self.stop_recording)
        self.stop_button.setEnabled(False)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addStretch()
        button_layout.addWidget(self.cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def start_recording(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Error", "Please enter a macro name!")
            return

        self.recording_macro = self.automation_manager.start_recording(name)
        self.recording_macro.description = self.description_edit.text().strip()

        self.status_label.setText("🔴 Recording in progress...")
        self.status_label.setStyleSheet("QLabel { color: red; font-weight: bold; }")

        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.name_edit.setEnabled(False)
        self.description_edit.setEnabled(False)

    def stop_recording(self):
        self.automation_manager.stop_recording()

        self.status_label.setText("✓ Recording completed")
        self.status_label.setStyleSheet("QLabel { color: green; font-weight: bold; }")

        QMessageBox.information(self, "Success",
            f"Macro '{self.recording_macro.name}' recorded successfully!\n"
            f"Actions recorded: {len(self.recording_macro.actions)}")

        self.accept()


class EditMacroDialog(QDialog):
    """Dialog for editing a macro script."""

    def __init__(self, automation_manager: AutomationManager, macro: Macro = None, parent=None):
        super().__init__(parent)
        self.automation_manager = automation_manager
        self.macro = macro
        self.setWindowTitle("Edit Macro" if macro else "New Macro")
        self.setModal(True)
        self.setMinimumSize(600, 500)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Basic info
        form_layout = QFormLayout()

        self.name_edit = QLineEdit()
        self.name_edit.setText(self.macro.name if self.macro else "")

        self.type_combo = QComboBox()
        self.type_combo.addItem("Scripted (Python)", MacroType.SCRIPTED)
        self.type_combo.addItem("Recorded Actions", MacroType.RECORDED)

        if self.macro:
            for i in range(self.type_combo.count()):
                if self.type_combo.itemData(i) == self.macro.macro_type:
                    self.type_combo.setCurrentIndex(i)

        self.enabled_check = QCheckBox("Enabled")
        self.enabled_check.setChecked(self.macro.enabled if self.macro else True)

        form_layout.addRow("Name:", self.name_edit)
        form_layout.addRow("Type:", self.type_combo)
        form_layout.addRow("", self.enabled_check)

        layout.addLayout(form_layout)

        # Script editor
        layout.addWidget(QLabel("Python Script:"))

        self.script_edit = QTextEdit()
        self.script_edit.setFontFamily("Courier")
        self.script_edit.setPlainText(self.macro.script if self.macro else
            "# Enter your Python script here\n"
            "# Available: parent (document object)\n\n"
            "# Example:\n"
            "# cursor = parent.text_edit.textCursor()\n"
            "# cursor.insertText('Hello, World!')\n")

        layout.addWidget(self.script_edit)

        # Keyboard shortcut
        shortcut_layout = QHBoxLayout()
        shortcut_layout.addWidget(QLabel("Keyboard Shortcut:"))

        self.shortcut_edit = QLineEdit()
        self.shortcut_edit.setPlaceholderText("e.g., Ctrl+Shift+M")
        self.shortcut_edit.setText(self.macro.keyboard_shortcut if self.macro else "")

        shortcut_layout.addWidget(self.shortcut_edit)
        shortcut_layout.addStretch()

        layout.addLayout(shortcut_layout)

        # Buttons
        button_layout = QHBoxLayout()

        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(self.accept)

        self.test_button = QPushButton("Test")
        self.test_button.clicked.connect(self.test_macro)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.test_button)
        button_layout.addStretch()
        button_layout.addWidget(self.cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def test_macro(self):
        # Create a temporary macro and test it
        test_macro = Macro("Test", MacroType.SCRIPTED)
        test_macro.script = self.script_edit.toPlainText()

        try:
            safe_globals = {'parent': self.automation_manager.parent, '__builtins__': {}}
            exec(test_macro.script, safe_globals)
            QMessageBox.information(self, "Success", "Macro executed successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Macro execution failed:\n{str(e)}")

    def accept(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Error", "Please enter a macro name!")
            return

        if not self.macro:
            macro_type = self.type_combo.currentData()
            self.macro = self.automation_manager.create_macro(name, macro_type)
        else:
            self.macro.name = name

        self.macro.enabled = self.enabled_check.isChecked()
        self.macro.script = self.script_edit.toPlainText()
        self.macro.keyboard_shortcut = self.shortcut_edit.text().strip()

        QMessageBox.information(self, "Success", "Macro saved successfully!")
        super().accept()


class AddInManagerDialog(QDialog):
    """Dialog for managing add-ins."""

    def __init__(self, automation_manager: AutomationManager, parent=None):
        super().__init__(parent)
        self.automation_manager = automation_manager
        self.setWindowTitle("Add-In Manager")
        self.setModal(True)
        self.setMinimumSize(700, 500)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Add-in table
        self.addon_table = QTableWidget()
        self.addon_table.setColumnCount(5)
        self.addon_table.setHorizontalHeaderLabels(
            ["Name", "Version", "Author", "Status", "Loaded"])
        self.addon_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.addon_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        self.refresh_addon_list()

        layout.addWidget(self.addon_table)

        # Details
        details_group = QGroupBox("Details")
        details_layout = QVBoxLayout()

        self.details_text = QTextEdit()
        self.details_text.setReadOnly(True)
        self.details_text.setMaximumHeight(100)

        details_layout.addWidget(self.details_text)
        details_group.setLayout(details_layout)

        layout.addWidget(details_group)

        # Buttons
        button_layout = QHBoxLayout()

        self.install_button = QPushButton("Install")
        self.install_button.clicked.connect(self.install_addon)

        self.enable_button = QPushButton("Enable")
        self.enable_button.clicked.connect(self.enable_addon)

        self.disable_button = QPushButton("Disable")
        self.disable_button.clicked.connect(self.disable_addon)

        self.uninstall_button = QPushButton("Uninstall")
        self.uninstall_button.clicked.connect(self.uninstall_addon)

        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.accept)

        button_layout.addWidget(self.install_button)
        button_layout.addWidget(self.enable_button)
        button_layout.addWidget(self.disable_button)
        button_layout.addWidget(self.uninstall_button)
        button_layout.addStretch()
        button_layout.addWidget(self.close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

        # Connect table selection
        self.addon_table.currentItemChanged.connect(self.on_addon_selected)

    def refresh_addon_list(self):
        self.addon_table.setRowCount(0)
        addons = self.automation_manager.get_all_addins()

        for addon in addons:
            row = self.addon_table.rowCount()
            self.addon_table.insertRow(row)

            self.addon_table.setItem(row, 0, QTableWidgetItem(addon.name))
            self.addon_table.setItem(row, 1, QTableWidgetItem(addon.version))
            self.addon_table.setItem(row, 2, QTableWidgetItem(addon.author))

            status = "Enabled" if addon.enabled else "Disabled"
            self.addon_table.setItem(row, 3, QTableWidgetItem(status))

            loaded = "Yes" if addon.loaded else "No"
            self.addon_table.setItem(row, 4, QTableWidgetItem(loaded))

            # Store add-in ID
            self.addon_table.item(row, 0).setData(Qt.ItemDataRole.UserRole, addon.id)

    def on_addon_selected(self, current, previous):
        if not current:
            return

        addon_id = self.addon_table.item(current.row(), 0).data(Qt.ItemDataRole.UserRole)
        addon = self.automation_manager.get_addon(addon_id)

        if addon:
            details = f"Name: {addon.name}\n"
            details += f"Version: {addon.version}\n"
            details += f"Author: {addon.author}\n"
            details += f"Description: {addon.description}\n"
            details += f"Installed: {addon.install_date.strftime('%Y-%m-%d')}\n"
            details += f"Dependencies: {', '.join(addon.dependencies) if addon.dependencies else 'None'}\n"
            self.details_text.setPlainText(details)

    def install_addon(self):
        dialog = InstallAddInDialog(self.automation_manager, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_addon_list()

    def enable_addon(self):
        row = self.addon_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Error", "Please select an add-in!")
            return

        addon_id = self.addon_table.item(row, 0).data(Qt.ItemDataRole.UserRole)

        if self.automation_manager.enable_addon(addon_id):
            QMessageBox.information(self, "Success", "Add-in enabled successfully!")
            self.refresh_addon_list()
        else:
            QMessageBox.critical(self, "Error", "Failed to enable add-in!")

    def disable_addon(self):
        row = self.addon_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Error", "Please select an add-in!")
            return

        addon_id = self.addon_table.item(row, 0).data(Qt.ItemDataRole.UserRole)

        if self.automation_manager.disable_addon(addon_id):
            QMessageBox.information(self, "Success", "Add-in disabled successfully!")
            self.refresh_addon_list()
        else:
            QMessageBox.critical(self, "Error", "Failed to disable add-in!")

    def uninstall_addon(self):
        row = self.addon_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Error", "Please select an add-in!")
            return

        reply = QMessageBox.question(self, "Confirm",
            "Are you sure you want to uninstall this add-in?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            addon_id = self.addon_table.item(row, 0).data(Qt.ItemDataRole.UserRole)
            if self.automation_manager.uninstall_addon(addon_id):
                QMessageBox.information(self, "Success", "Add-in uninstalled successfully!")
                self.refresh_addon_list()
            else:
                QMessageBox.critical(self, "Error", "Failed to uninstall add-in!")


class InstallAddInDialog(QDialog):
    """Dialog for installing an add-in."""

    def __init__(self, automation_manager: AutomationManager, parent=None):
        super().__init__(parent)
        self.automation_manager = automation_manager
        self.setWindowTitle("Install Add-In")
        self.setModal(True)
        self.setMinimumWidth(500)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        form_layout = QFormLayout()

        self.name_edit = QLineEdit()
        self.version_edit = QLineEdit()
        self.version_edit.setText("1.0.0")
        self.author_edit = QLineEdit()

        self.description_edit = QTextEdit()
        self.description_edit.setMaximumHeight(60)

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)

        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self.browse_file)

        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(self.browse_button)

        form_layout.addRow("Name:", self.name_edit)
        form_layout.addRow("Version:", self.version_edit)
        form_layout.addRow("Author:", self.author_edit)
        form_layout.addRow("Description:", self.description_edit)
        form_layout.addRow("File:", file_layout)

        layout.addLayout(form_layout)

        # Buttons
        button_layout = QHBoxLayout()

        self.install_button = QPushButton("Install")
        self.install_button.clicked.connect(self.accept)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(self.install_button)
        button_layout.addWidget(self.cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Add-In File", "", "Python Files (*.py);;All Files (*)")

        if file_path:
            self.file_path_edit.setText(file_path)

    def accept(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Error", "Add-in name is required!")
            return

        addon = AddIn(name, self.version_edit.text().strip())
        addon.author = self.author_edit.text().strip()
        addon.description = self.description_edit.toPlainText().strip()
        addon.file_path = self.file_path_edit.text()

        if self.automation_manager.install_addon(addon):
            QMessageBox.information(self, "Success", "Add-in installed successfully!")
            super().accept()
        else:
            QMessageBox.critical(self, "Error", "Failed to install add-in!")


class CustomXMLDialog(QDialog):
    """Dialog for managing custom XML data."""

    def __init__(self, automation_manager: AutomationManager, parent=None):
        super().__init__(parent)
        self.automation_manager = automation_manager
        self.setWindowTitle("Custom XML Data")
        self.setModal(True)
        self.setMinimumSize(800, 600)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Split view: list and editor
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # XML parts list
        list_widget = QWidget()
        list_layout = QVBoxLayout()

        list_layout.addWidget(QLabel("XML Parts:"))

        self.xml_list = QListWidget()
        self.xml_list.currentItemChanged.connect(self.on_xml_part_selected)
        self.refresh_xml_list()

        list_layout.addWidget(self.xml_list)

        # List buttons
        list_button_layout = QHBoxLayout()

        self.add_xml_button = QPushButton("Add")
        self.add_xml_button.clicked.connect(self.add_xml_part)

        self.delete_xml_button = QPushButton("Delete")
        self.delete_xml_button.clicked.connect(self.delete_xml_part)

        list_button_layout.addWidget(self.add_xml_button)
        list_button_layout.addWidget(self.delete_xml_button)

        list_layout.addLayout(list_button_layout)

        list_widget.setLayout(list_layout)
        splitter.addWidget(list_widget)

        # XML editor
        editor_widget = QWidget()
        editor_layout = QVBoxLayout()

        # Part info
        info_layout = QFormLayout()

        self.part_name_label = QLabel("-")
        self.part_namespace_label = QLabel("-")

        info_layout.addRow("Name:", self.part_name_label)
        info_layout.addRow("Namespace:", self.part_namespace_label)

        editor_layout.addLayout(info_layout)

        # XML content editor
        editor_layout.addWidget(QLabel("XML Content:"))

        self.xml_editor = QTextEdit()
        self.xml_editor.setFontFamily("Courier")
        self.xml_editor.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)

        editor_layout.addWidget(self.xml_editor)

        # Editor buttons
        editor_button_layout = QHBoxLayout()

        self.save_xml_button = QPushButton("Save Changes")
        self.save_xml_button.clicked.connect(self.save_xml_content)

        self.validate_button = QPushButton("Validate XML")
        self.validate_button.clicked.connect(self.validate_xml)

        self.import_button = QPushButton("Import from File")
        self.import_button.clicked.connect(self.import_xml)

        self.export_button = QPushButton("Export to File")
        self.export_button.clicked.connect(self.export_xml)

        editor_button_layout.addWidget(self.save_xml_button)
        editor_button_layout.addWidget(self.validate_button)
        editor_button_layout.addWidget(self.import_button)
        editor_button_layout.addWidget(self.export_button)

        editor_layout.addLayout(editor_button_layout)

        editor_widget.setLayout(editor_layout)
        splitter.addWidget(editor_widget)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        layout.addWidget(splitter)

        # Dialog buttons
        button_layout = QHBoxLayout()

        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.accept)

        button_layout.addStretch()
        button_layout.addWidget(self.close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

        self.current_xml_part = None

    def refresh_xml_list(self):
        self.xml_list.clear()
        xml_parts = self.automation_manager.get_all_xml_parts()

        for part in xml_parts:
            item = QListWidgetItem(part.name)
            item.setData(Qt.ItemDataRole.UserRole, part.id)
            self.xml_list.addItem(item)

    def on_xml_part_selected(self, current, previous):
        if not current:
            return

        part_id = current.data(Qt.ItemDataRole.UserRole)
        self.current_xml_part = self.automation_manager.get_xml_part(part_id)

        if self.current_xml_part:
            self.part_name_label.setText(self.current_xml_part.name)
            self.part_namespace_label.setText(self.current_xml_part.namespace)
            self.xml_editor.setPlainText(self.current_xml_part.get_xml_content())

    def add_xml_part(self):
        from PySide6.QtWidgets import QInputDialog

        name, ok = QInputDialog.getText(self, "New XML Part", "Enter XML part name:")
        if ok and name:
            namespace, ok = QInputDialog.getText(self, "Namespace",
                "Enter namespace (optional):", text=f"http://pyword.example.com/{name}")

            xml_part = self.automation_manager.create_xml_part(name, namespace if ok else "")
            self.refresh_xml_list()
            QMessageBox.information(self, "Success", "XML part created successfully!")

    def delete_xml_part(self):
        item = self.xml_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Error", "Please select an XML part to delete!")
            return

        reply = QMessageBox.question(self, "Confirm",
            "Are you sure you want to delete this XML part?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            part_id = item.data(Qt.ItemDataRole.UserRole)
            self.automation_manager.delete_xml_part(part_id)
            self.refresh_xml_list()
            self.xml_editor.clear()

    def save_xml_content(self):
        if not self.current_xml_part:
            QMessageBox.warning(self, "Error", "No XML part selected!")
            return

        xml_content = self.xml_editor.toPlainText()

        if self.current_xml_part.set_xml_content(xml_content):
            QMessageBox.information(self, "Success", "XML content saved successfully!")
        else:
            QMessageBox.critical(self, "Error", "Invalid XML content! Please check and try again.")

    def validate_xml(self):
        xml_content = self.xml_editor.toPlainText()

        import xml.etree.ElementTree as ET
        try:
            ET.fromstring(xml_content)
            QMessageBox.information(self, "Validation", "XML is valid!")
        except ET.ParseError as e:
            QMessageBox.critical(self, "Validation Error", f"XML is invalid:\n{str(e)}")

    def import_xml(self):
        if not self.current_xml_part:
            QMessageBox.warning(self, "Error", "Please select an XML part first!")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Import XML", "", "XML Files (*.xml);;All Files (*)")

        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    xml_content = f.read()

                if self.current_xml_part.set_xml_content(xml_content):
                    self.xml_editor.setPlainText(xml_content)
                    QMessageBox.information(self, "Success", "XML imported successfully!")
                else:
                    QMessageBox.critical(self, "Error", "Invalid XML content!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to import XML:\n{str(e)}")

    def export_xml(self):
        if not self.current_xml_part:
            QMessageBox.warning(self, "Error", "No XML part selected!")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export XML", f"{self.current_xml_part.name}.xml",
            "XML Files (*.xml);;All Files (*)")

        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.current_xml_part.get_xml_content())
                QMessageBox.information(self, "Success", "XML exported successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to export XML:\n{str(e)}")
