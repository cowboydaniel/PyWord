"""
Mail Merge System for PyWord.

This module provides comprehensive mail merge functionality, including
data source integration, field insertion, preview, and document generation.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QListWidget, QListWidgetItem, QGroupBox,
                               QComboBox, QTextEdit, QFormLayout, QMessageBox,
                               QCheckBox, QSpinBox, QTableWidget, QTableWidgetItem,
                               QHeaderView, QFileDialog, QTabWidget, QWidget,
                               QRadioButton, QButtonGroup, QProgressDialog, QWizard,
                               QWizardPage)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QTextCursor, QTextCharFormat, QFont, QColor
from datetime import datetime
import json
import csv
import uuid


class DataSource:
    """Represents a data source for mail merge."""

    def __init__(self, name, source_type='csv'):
        self.id = str(uuid.uuid4())
        self.name = name
        self.source_type = source_type  # 'csv', 'json', 'manual'
        self.file_path = None
        self.records = []  # List of dictionaries
        self.field_names = []  # List of field names
        self.created = datetime.now()

    def load_from_csv(self, file_path):
        """Load data from CSV file."""
        try:
            self.records = []
            self.field_names = []

            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                self.field_names = reader.fieldnames or []

                for row in reader:
                    self.records.append(row)

            self.file_path = file_path
            return True
        except Exception as e:
            print(f"Error loading CSV: {e}")
            return False

    def load_from_json(self, file_path):
        """Load data from JSON file."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            if isinstance(data, list) and len(data) > 0:
                self.records = data
                # Get field names from first record
                if isinstance(data[0], dict):
                    self.field_names = list(data[0].keys())
                else:
                    return False

                self.file_path = file_path
                return True
            return False
        except Exception as e:
            print(f"Error loading JSON: {e}")
            return False

    def add_manual_record(self, record):
        """Add a record manually."""
        # Update field names if needed
        for key in record.keys():
            if key not in self.field_names:
                self.field_names.append(key)

        self.records.append(record)

    def get_record(self, index):
        """Get a record by index."""
        if 0 <= index < len(self.records):
            return self.records[index]
        return None

    def get_field_value(self, record_index, field_name):
        """Get a specific field value from a record."""
        record = self.get_record(record_index)
        if record:
            return record.get(field_name, '')
        return ''

    def to_dict(self):
        """Convert to dictionary for serialization."""
        return {
            'id': self.id,
            'name': self.name,
            'source_type': self.source_type,
            'file_path': self.file_path,
            'records': self.records,
            'field_names': self.field_names,
            'created': self.created.isoformat()
        }

    @staticmethod
    def from_dict(data):
        """Create DataSource from dictionary."""
        ds = DataSource(data['name'], data.get('source_type', 'csv'))
        ds.id = data['id']
        ds.file_path = data.get('file_path')
        ds.records = data.get('records', [])
        ds.field_names = data.get('field_names', [])
        if 'created' in data:
            ds.created = datetime.fromisoformat(data['created'])
        return ds


class MergeField:
    """Represents a merge field in the document."""

    def __init__(self, field_name, position=0, format_string=None):
        self.id = str(uuid.uuid4())
        self.field_name = field_name
        self.position = position
        self.format_string = format_string  # Optional formatting (e.g., uppercase, date format)

    def to_dict(self):
        """Convert to dictionary for serialization."""
        return {
            'id': self.id,
            'field_name': self.field_name,
            'position': self.position,
            'format_string': self.format_string
        }

    @staticmethod
    def from_dict(data):
        """Create MergeField from dictionary."""
        field = MergeField(
            data['field_name'],
            data.get('position', 0),
            data.get('format_string')
        )
        field.id = data['id']
        return field


class MailMergeManager:
    """Manages mail merge operations."""

    def __init__(self, parent):
        self.parent = parent  # Reference to text editor
        self.data_sources = []
        self.active_data_source = None
        self.merge_fields = []

        # Settings
        self.field_delimiter = '«»'  # Delimiters for merge fields
        self.preview_record_index = 0

    def add_data_source(self, data_source):
        """Add a data source."""
        self.data_sources.append(data_source)
        if not self.active_data_source:
            self.active_data_source = data_source
        return data_source

    def set_active_data_source(self, data_source_id):
        """Set the active data source."""
        for ds in self.data_sources:
            if ds.id == data_source_id:
                self.active_data_source = ds
                return True
        return False

    def get_data_source_by_id(self, data_source_id):
        """Get a data source by ID."""
        for ds in self.data_sources:
            if ds.id == data_source_id:
                return ds
        return None

    def insert_merge_field(self, field_name, position=None):
        """Insert a merge field in the document."""
        if not self.active_data_source:
            QMessageBox.warning(
                None,
                "No Data Source",
                "Please select a data source first."
            )
            return None

        if field_name not in self.active_data_source.field_names:
            QMessageBox.warning(
                None,
                "Invalid Field",
                f"Field '{field_name}' not found in data source."
            )
            return None

        if position is None:
            cursor = self.parent.textCursor()
            position = cursor.position()
        else:
            cursor = QTextCursor(self.parent.document())
            cursor.setPosition(position)

        # Create merge field
        merge_field = MergeField(field_name, position)
        self.merge_fields.append(merge_field)

        # Insert field marker in document
        self._insert_field_marker(cursor, field_name)

        return merge_field

    def _insert_field_marker(self, cursor, field_name):
        """Insert a visual marker for a merge field."""
        # Format for merge field
        char_format = QTextCharFormat()
        char_format.setBackground(QColor(220, 220, 220))  # Light gray background
        char_format.setForeground(QColor(0, 0, 255))  # Blue text
        char_format.setFontItalic(True)

        field_text = f"{self.field_delimiter[0]}{field_name}{self.field_delimiter[1]}"

        cursor.setCharFormat(char_format)
        cursor.insertText(field_text)

        # Reset format
        cursor.setCharFormat(QTextCharFormat())

    def preview_merge(self, record_index=None):
        """Preview the merge with a specific record."""
        if not self.active_data_source:
            return False

        if record_index is None:
            record_index = self.preview_record_index

        if record_index >= len(self.active_data_source.records):
            return False

        # Get the document text
        document = self.parent.document()
        original_text = document.toPlainText()

        # Replace merge fields with actual data
        merged_text = self._replace_merge_fields(original_text, record_index)

        # Update document (temporarily for preview)
        # In a real implementation, you'd use a separate preview window
        # to avoid modifying the original document

        return True

    def _replace_merge_fields(self, text, record_index):
        """Replace merge fields in text with actual data."""
        if not self.active_data_source:
            return text

        record = self.active_data_source.get_record(record_index)
        if not record:
            return text

        result = text

        # Replace each field
        for field_name in self.active_data_source.field_names:
            field_marker = f"{self.field_delimiter[0]}{field_name}{self.field_delimiter[1]}"
            field_value = record.get(field_name, '')

            result = result.replace(field_marker, str(field_value))

        return result

    def complete_merge(self, output_type='separate_documents'):
        """Complete the mail merge operation."""
        if not self.active_data_source:
            QMessageBox.warning(
                None,
                "No Data Source",
                "Please select a data source first."
            )
            return []

        if not self.active_data_source.records:
            QMessageBox.warning(
                None,
                "No Records",
                "Data source has no records."
            )
            return []

        # Get template document
        template_text = self.parent.document().toPlainText()

        merged_documents = []

        # Progress dialog
        progress = QProgressDialog(
            "Merging documents...",
            "Cancel",
            0,
            len(self.active_data_source.records)
        )
        progress.setWindowModality(Qt.WindowModality.WindowModal)

        for i, record in enumerate(self.active_data_source.records):
            if progress.wasCanceled():
                break

            # Replace merge fields
            merged_text = self._replace_merge_fields(template_text, i)
            merged_documents.append(merged_text)

            progress.setValue(i + 1)

        progress.close()

        return merged_documents

    def save_merged_documents(self, documents, output_dir, base_name='merged_doc'):
        """Save merged documents to files."""
        saved_files = []

        for i, doc_text in enumerate(documents):
            file_path = f"{output_dir}/{base_name}_{i+1}.txt"

            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(doc_text)
                saved_files.append(file_path)
            except Exception as e:
                print(f"Error saving document {i+1}: {e}")

        return saved_files

    def get_merge_fields_in_document(self):
        """Get all merge field markers in the document."""
        text = self.parent.document().toPlainText()
        fields = []

        # Find all field markers
        import re
        pattern = f"{re.escape(self.field_delimiter[0])}(.*?){re.escape(self.field_delimiter[1])}"
        matches = re.findall(pattern, text)

        return list(set(matches))  # Return unique field names

    def validate_merge_fields(self):
        """Validate that all merge fields in document exist in data source."""
        if not self.active_data_source:
            return False, ["No data source selected"]

        fields_in_doc = self.get_merge_fields_in_document()
        available_fields = self.active_data_source.field_names

        invalid_fields = []
        for field in fields_in_doc:
            if field not in available_fields:
                invalid_fields.append(field)

        if invalid_fields:
            return False, invalid_fields

        return True, []


class DataSourceDialog(QDialog):
    """Dialog for creating/editing a data source."""

    def __init__(self, data_source=None, parent=None):
        super().__init__(parent)
        self.data_source = data_source if data_source else DataSource("New Data Source")

        self.setWindowTitle("Data Source")
        self.setModal(True)
        self.setMinimumSize(700, 500)

        self.setup_ui()

        if data_source:
            self.load_data_source()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Data source name
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Name:"))
        self.name_edit = QLineEdit()
        self.name_edit.setText(self.data_source.name)
        name_layout.addWidget(self.name_edit)
        layout.addLayout(name_layout)

        # Source type
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Source type:"))

        self.type_combo = QComboBox()
        self.type_combo.addItems(['CSV File', 'JSON File', 'Manual Entry'])
        self.type_combo.currentTextChanged.connect(self.on_type_changed)
        type_layout.addWidget(self.type_combo)
        type_layout.addStretch()

        layout.addLayout(type_layout)

        # File selection (for CSV/JSON)
        self.file_group = QGroupBox("File Selection")
        file_layout = QHBoxLayout()

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        file_layout.addWidget(self.file_path_edit)

        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_button)

        load_button = QPushButton("Load")
        load_button.clicked.connect(self.load_file)
        file_layout.addWidget(load_button)

        self.file_group.setLayout(file_layout)
        layout.addWidget(self.file_group)

        # Data table
        table_label = QLabel("<b>Data Preview:</b>")
        layout.addWidget(table_label)

        self.data_table = QTableWidget()
        self.data_table.setAlternatingRowColors(True)
        layout.addWidget(self.data_table)

        # Manual entry buttons (for manual entry mode)
        self.manual_group = QGroupBox("Manual Entry")
        manual_layout = QHBoxLayout()

        add_record_button = QPushButton("Add Record")
        add_record_button.clicked.connect(self.add_manual_record)
        manual_layout.addWidget(add_record_button)

        edit_record_button = QPushButton("Edit Record")
        edit_record_button.clicked.connect(self.edit_manual_record)
        manual_layout.addWidget(edit_record_button)

        delete_record_button = QPushButton("Delete Record")
        delete_record_button.clicked.connect(self.delete_manual_record)
        manual_layout.addWidget(delete_record_button)

        manual_layout.addStretch()

        self.manual_group.setLayout(manual_layout)
        self.manual_group.setVisible(False)
        layout.addWidget(self.manual_group)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.save_data_source)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def on_type_changed(self):
        """Handle source type change."""
        source_type = self.type_combo.currentText()

        if source_type == 'Manual Entry':
            self.file_group.setVisible(False)
            self.manual_group.setVisible(True)
        else:
            self.file_group.setVisible(True)
            self.manual_group.setVisible(False)

    def browse_file(self):
        """Browse for a data file."""
        source_type = self.type_combo.currentText()

        if source_type == 'CSV File':
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Select CSV File",
                "",
                "CSV Files (*.csv)"
            )
        elif source_type == 'JSON File':
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Select JSON File",
                "",
                "JSON Files (*.json)"
            )
        else:
            return

        if file_path:
            self.file_path_edit.setText(file_path)

    def load_file(self):
        """Load data from file."""
        file_path = self.file_path_edit.text()
        if not file_path:
            QMessageBox.warning(self, "No File", "Please select a file first.")
            return

        source_type = self.type_combo.currentText()

        success = False
        if source_type == 'CSV File':
            success = self.data_source.load_from_csv(file_path)
        elif source_type == 'JSON File':
            success = self.data_source.load_from_json(file_path)

        if success:
            self.refresh_table()
            QMessageBox.information(
                self,
                "Success",
                f"Loaded {len(self.data_source.records)} records."
            )
        else:
            QMessageBox.warning(self, "Error", "Failed to load file.")

    def refresh_table(self):
        """Refresh the data table."""
        self.data_table.clear()

        if not self.data_source.records:
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            return

        # Set up table
        self.data_table.setColumnCount(len(self.data_source.field_names))
        self.data_table.setHorizontalHeaderLabels(self.data_source.field_names)
        self.data_table.setRowCount(len(self.data_source.records))

        # Populate table
        for row_idx, record in enumerate(self.data_source.records):
            for col_idx, field_name in enumerate(self.data_source.field_names):
                value = record.get(field_name, '')
                item = QTableWidgetItem(str(value))
                self.data_table.setItem(row_idx, col_idx, item)

        # Resize columns to content
        self.data_table.resizeColumnsToContents()

    def add_manual_record(self):
        """Add a manual record."""
        # This would open a dialog to add a record
        # Simplified implementation
        QMessageBox.information(
            self,
            "Add Record",
            "Manual record entry would be implemented here."
        )

    def edit_manual_record(self):
        """Edit a manual record."""
        pass

    def delete_manual_record(self):
        """Delete a manual record."""
        pass

    def load_data_source(self):
        """Load existing data source into dialog."""
        self.name_edit.setText(self.data_source.name)

        type_map = {
            'csv': 'CSV File',
            'json': 'JSON File',
            'manual': 'Manual Entry'
        }
        self.type_combo.setCurrentText(type_map.get(self.data_source.source_type, 'CSV File'))

        if self.data_source.file_path:
            self.file_path_edit.setText(self.data_source.file_path)

        self.refresh_table()

    def save_data_source(self):
        """Save the data source."""
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Invalid Name", "Please enter a data source name.")
            return

        self.data_source.name = name

        type_map = {
            'CSV File': 'csv',
            'JSON File': 'json',
            'Manual Entry': 'manual'
        }
        self.data_source.source_type = type_map.get(
            self.type_combo.currentText(),
            'csv'
        )

        self.accept()

    def get_data_source(self):
        """Get the data source."""
        return self.data_source


class MailMergeWizard(QWizard):
    """Wizard for guiding through mail merge process."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager

        self.setWindowTitle("Mail Merge Wizard")
        self.setMinimumSize(700, 500)

        # Add pages
        self.addPage(self._create_select_data_page())
        self.addPage(self._create_insert_fields_page())
        self.addPage(self._create_preview_page())
        self.addPage(self._create_complete_page())

    def _create_select_data_page(self):
        """Create data source selection page."""
        page = QWizardPage()
        page.setTitle("Select Data Source")
        page.setSubTitle("Choose or create a data source for the mail merge.")

        layout = QVBoxLayout()

        # Data source list
        list_label = QLabel("<b>Available Data Sources:</b>")
        layout.addWidget(list_label)

        self.data_source_list = QListWidget()
        self.refresh_data_sources()
        layout.addWidget(self.data_source_list)

        # Buttons
        button_layout = QHBoxLayout()

        new_button = QPushButton("New Data Source...")
        new_button.clicked.connect(self.create_data_source)
        button_layout.addWidget(new_button)

        edit_button = QPushButton("Edit...")
        edit_button.clicked.connect(self.edit_data_source)
        button_layout.addWidget(edit_button)

        button_layout.addStretch()

        layout.addLayout(button_layout)
        page.setLayout(layout)

        return page

    def _create_insert_fields_page(self):
        """Create field insertion page."""
        page = QWizardPage()
        page.setTitle("Insert Merge Fields")
        page.setSubTitle("Insert merge fields into your document.")

        layout = QVBoxLayout()

        info_label = QLabel(
            "Select a field and click 'Insert Field' to add it to your document at the cursor position."
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        # Available fields
        fields_label = QLabel("<b>Available Fields:</b>")
        layout.addWidget(fields_label)

        self.fields_list = QListWidget()
        layout.addWidget(self.fields_list)

        # Insert button
        insert_button = QPushButton("Insert Field")
        insert_button.clicked.connect(self.insert_field)
        layout.addWidget(insert_button)

        page.setLayout(layout)

        return page

    def _create_preview_page(self):
        """Create preview page."""
        page = QWizardPage()
        page.setTitle("Preview Results")
        page.setSubTitle("Preview how the merge will look for each record.")

        layout = QVBoxLayout()

        # Navigation
        nav_layout = QHBoxLayout()
        nav_layout.addWidget(QLabel("Record:"))

        prev_button = QPushButton("< Previous")
        prev_button.clicked.connect(self.preview_previous)
        nav_layout.addWidget(prev_button)

        self.record_label = QLabel("1 of 1")
        nav_layout.addWidget(self.record_label)

        next_button = QPushButton("Next >")
        next_button.clicked.connect(self.preview_next)
        nav_layout.addWidget(next_button)

        nav_layout.addStretch()

        layout.addLayout(nav_layout)

        # Preview text
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        layout.addWidget(self.preview_text)

        page.setLayout(layout)

        return page

    def _create_complete_page(self):
        """Create completion page."""
        page = QWizardPage()
        page.setTitle("Complete Merge")
        page.setSubTitle("Choose how to output the merged documents.")

        layout = QVBoxLayout()

        # Output options
        options_group = QGroupBox("Output Options")
        options_layout = QVBoxLayout()

        self.output_group = QButtonGroup()

        self.separate_docs_radio = QRadioButton("Create separate documents")
        self.separate_docs_radio.setChecked(True)
        self.output_group.addButton(self.separate_docs_radio)
        options_layout.addWidget(self.separate_docs_radio)

        self.single_doc_radio = QRadioButton("Create single document")
        self.output_group.addButton(self.single_doc_radio)
        options_layout.addWidget(self.single_doc_radio)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Output directory
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(QLabel("Output directory:"))
        self.output_dir_edit = QLineEdit()
        dir_layout.addWidget(self.output_dir_edit)

        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self.browse_output_dir)
        dir_layout.addWidget(browse_button)

        layout.addLayout(dir_layout)

        layout.addStretch()

        # Merge button
        merge_button = QPushButton("Complete Merge")
        merge_button.clicked.connect(self.complete_merge)
        layout.addWidget(merge_button)

        page.setLayout(layout)

        return page

    def refresh_data_sources(self):
        """Refresh the data sources list."""
        self.data_source_list.clear()

        for ds in self.manager.data_sources:
            item_text = f"{ds.name} ({len(ds.records)} records)"
            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, ds.id)
            self.data_source_list.addItem(item)

    def create_data_source(self):
        """Create a new data source."""
        dialog = DataSourceDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            ds = dialog.get_data_source()
            self.manager.add_data_source(ds)
            self.refresh_data_sources()

    def edit_data_source(self):
        """Edit selected data source."""
        items = self.data_source_list.selectedItems()
        if not items:
            return

        ds_id = items[0].data(Qt.ItemDataRole.UserRole)
        ds = self.manager.get_data_source_by_id(ds_id)

        if ds:
            dialog = DataSourceDialog(ds, parent=self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.refresh_data_sources()

    def insert_field(self):
        """Insert selected field."""
        items = self.fields_list.selectedItems()
        if not items:
            return

        field_name = items[0].text()
        self.manager.insert_merge_field(field_name)

        QMessageBox.information(
            self,
            "Field Inserted",
            f"Field '{field_name}' has been inserted at the cursor position."
        )

    def preview_previous(self):
        """Preview previous record."""
        if self.manager.preview_record_index > 0:
            self.manager.preview_record_index -= 1
            self.update_preview()

    def preview_next(self):
        """Preview next record."""
        if self.manager.active_data_source:
            if self.manager.preview_record_index < len(self.manager.active_data_source.records) - 1:
                self.manager.preview_record_index += 1
                self.update_preview()

    def update_preview(self):
        """Update the preview display."""
        if not self.manager.active_data_source:
            return

        total_records = len(self.manager.active_data_source.records)
        current_index = self.manager.preview_record_index

        self.record_label.setText(f"{current_index + 1} of {total_records}")

        # Get template and merge
        template_text = self.manager.parent.document().toPlainText()
        merged_text = self.manager._replace_merge_fields(template_text, current_index)

        self.preview_text.setPlainText(merged_text)

    def browse_output_dir(self):
        """Browse for output directory."""
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory"
        )

        if dir_path:
            self.output_dir_edit.setText(dir_path)

    def complete_merge(self):
        """Complete the mail merge."""
        output_dir = self.output_dir_edit.text()
        if not output_dir:
            QMessageBox.warning(self, "No Directory", "Please select an output directory.")
            return

        # Complete merge
        merged_docs = self.manager.complete_merge()

        if merged_docs:
            # Save documents
            saved_files = self.manager.save_merged_documents(merged_docs, output_dir)

            QMessageBox.information(
                self,
                "Merge Complete",
                f"Successfully created {len(saved_files)} merged documents in:\n{output_dir}"
            )
        else:
            QMessageBox.warning(self, "Error", "Failed to complete merge.")


class MailMergeToolbar(QWidget):
    """Toolbar widget for mail merge operations."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager

        self.setup_ui()

    def setup_ui(self):
        """Setup the toolbar UI."""
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        layout.addWidget(QLabel("Mail Merge:"))

        # Start wizard button
        wizard_button = QPushButton("Start Mail Merge...")
        wizard_button.clicked.connect(self.start_wizard)
        layout.addWidget(wizard_button)

        # Quick insert field
        layout.addWidget(QLabel("Insert:"))
        self.field_combo = QComboBox()
        self.field_combo.currentTextChanged.connect(self.on_field_selected)
        layout.addWidget(self.field_combo)

        layout.addStretch()

        self.setLayout(layout)

    def start_wizard(self):
        """Start the mail merge wizard."""
        wizard = MailMergeWizard(self.manager, parent=self)
        wizard.exec()

    def on_field_selected(self, field_name):
        """Handle field selection."""
        if field_name and field_name != "Select field...":
            self.manager.insert_merge_field(field_name)
            # Reset combo
            self.field_combo.setCurrentIndex(0)

    def refresh_fields(self):
        """Refresh the fields combo box."""
        self.field_combo.clear()
        self.field_combo.addItem("Select field...")

        if self.manager.active_data_source:
            self.field_combo.addItems(self.manager.active_data_source.field_names)
