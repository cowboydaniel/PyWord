"""
Document Comparison for PyWord.

This module provides functionality to compare two documents and highlight
the differences between them.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit,
                               QPushButton, QFileDialog, QMessageBox, QSplitter,
                               QWidget, QGroupBox, QRadioButton, QButtonGroup,
                               QCheckBox, QComboBox, QListWidget, QListWidgetItem,
                               QTabWidget)
from PySide6.QtCore import Qt
from PySide6.QtGui import QTextCursor, QTextCharFormat, QColor, QFont, QTextDocument
import difflib
from datetime import datetime


class DifferenceType:
    """Types of differences between documents."""
    ADDITION = "Addition"
    DELETION = "Deletion"
    MODIFICATION = "Modification"
    NO_CHANGE = "No Change"


class Difference:
    """Represents a difference between two documents."""

    def __init__(self, diff_type, position, original_text, new_text):
        self.type = diff_type
        self.position = position
        self.original_text = original_text
        self.new_text = new_text


class DocumentComparison:
    """Manages document comparison."""

    def __init__(self, original_doc, revised_doc):
        self.original_doc = original_doc
        self.revised_doc = revised_doc
        self.differences = []

        # Formatting for differences
        self.addition_format = QTextCharFormat()
        self.addition_format.setBackground(QColor(200, 255, 200))  # Light green
        self.addition_format.setForeground(QColor(0, 128, 0))

        self.deletion_format = QTextCharFormat()
        self.deletion_format.setBackground(QColor(255, 200, 200))  # Light red
        self.deletion_format.setForeground(QColor(255, 0, 0))
        self.deletion_format.setFontStrikeOut(True)

        self.modification_format = QTextCharFormat()
        self.modification_format.setBackground(QColor(255, 255, 200))  # Light yellow
        self.modification_format.setForeground(QColor(128, 128, 0))

    def compare(self):
        """Compare the two documents and generate differences."""
        self.differences.clear()

        # Get text from both documents
        original_text = self.original_doc.toPlainText()
        revised_text = self.revised_doc.toPlainText()

        # Split into lines for comparison
        original_lines = original_text.splitlines(keepends=True)
        revised_lines = revised_text.splitlines(keepends=True)

        # Use difflib to find differences
        differ = difflib.Differ()
        diff = list(differ.compare(original_lines, revised_lines))

        position = 0
        for line in diff:
            if line.startswith('- '):
                # Deletion
                text = line[2:]
                diff = Difference(DifferenceType.DELETION, position, text, "")
                self.differences.append(diff)
            elif line.startswith('+ '):
                # Addition
                text = line[2:]
                diff = Difference(DifferenceType.ADDITION, position, "", text)
                self.differences.append(diff)
            elif line.startswith('? '):
                # Modification marker (skip for now)
                continue
            else:
                # No change
                position += len(line[2:])

        return self.differences

    def compare_detailed(self):
        """Compare documents with detailed word-level differences."""
        self.differences.clear()

        original_text = self.original_doc.toPlainText()
        revised_text = self.revised_doc.toPlainText()

        # Split into words
        original_words = original_text.split()
        revised_words = revised_text.split()

        # Use SequenceMatcher for more detailed comparison
        matcher = difflib.SequenceMatcher(None, original_words, revised_words)

        position = 0
        for op, i1, i2, j1, j2 in matcher.get_opcodes():
            if op == 'delete':
                text = ' '.join(original_words[i1:i2])
                diff = Difference(DifferenceType.DELETION, position, text, "")
                self.differences.append(diff)
            elif op == 'insert':
                text = ' '.join(revised_words[j1:j2])
                diff = Difference(DifferenceType.ADDITION, position, "", text)
                self.differences.append(diff)
            elif op == 'replace':
                orig_text = ' '.join(original_words[i1:i2])
                new_text = ' '.join(revised_words[j1:j2])
                diff = Difference(DifferenceType.MODIFICATION, position, orig_text, new_text)
                self.differences.append(diff)
            # 'equal' means no change, we can skip

            # Update position
            if op != 'delete':
                position += sum(len(w) + 1 for w in revised_words[j1:j2])

        return self.differences

    def get_summary(self):
        """Get a summary of differences."""
        additions = sum(1 for d in self.differences if d.type == DifferenceType.ADDITION)
        deletions = sum(1 for d in self.differences if d.type == DifferenceType.DELETION)
        modifications = sum(1 for d in self.differences if d.type == DifferenceType.MODIFICATION)

        return {
            'additions': additions,
            'deletions': deletions,
            'modifications': modifications,
            'total': len(self.differences)
        }

    def apply_highlighting(self, text_edit, show_original=True):
        """Apply highlighting to show differences in a text edit widget."""
        cursor = text_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.Start)

        for diff in self.differences:
            if diff.type == DifferenceType.ADDITION:
                cursor.insertText(diff.new_text, self.addition_format)
            elif diff.type == DifferenceType.DELETION:
                if show_original:
                    cursor.insertText(diff.original_text, self.deletion_format)
            elif diff.type == DifferenceType.MODIFICATION:
                if show_original:
                    cursor.insertText(diff.original_text, self.deletion_format)
                cursor.insertText(" → ", self.modification_format)
                cursor.insertText(diff.new_text, self.addition_format)


class DocumentComparisonDialog(QDialog):
    """Dialog for comparing two documents."""

    def __init__(self, current_document, parent=None):
        super().__init__(parent)
        self.current_document = current_document
        self.comparison = None
        self.setWindowTitle("Compare Documents")
        self.setMinimumSize(900, 700)

        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Header
        header = QLabel("<h2>Document Comparison</h2>")
        layout.addWidget(header)

        # File selection
        file_group = QGroupBox("Select Document to Compare")
        file_layout = QHBoxLayout()

        self.file_label = QLabel("No file selected")
        file_layout.addWidget(self.file_label)

        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_button)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Comparison options
        options_group = QGroupBox("Comparison Options")
        options_layout = QVBoxLayout()

        self.comparison_type_group = QButtonGroup()
        self.line_compare_radio = QRadioButton("Line-by-line comparison")
        self.word_compare_radio = QRadioButton("Word-by-word comparison (detailed)")

        self.comparison_type_group.addButton(self.line_compare_radio)
        self.comparison_type_group.addButton(self.word_compare_radio)

        self.line_compare_radio.setChecked(True)

        options_layout.addWidget(self.line_compare_radio)
        options_layout.addWidget(self.word_compare_radio)

        self.show_deletions_checkbox = QCheckBox("Show deletions")
        self.show_deletions_checkbox.setChecked(True)
        options_layout.addWidget(self.show_deletions_checkbox)

        compare_button = QPushButton("Compare Documents")
        compare_button.clicked.connect(self.compare_documents)
        options_layout.addWidget(compare_button)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Results tabs
        self.results_tabs = QTabWidget()

        # Summary tab
        summary_tab = self.create_summary_tab()
        self.results_tabs.addTab(summary_tab, "Summary")

        # Side-by-side view tab
        side_by_side_tab = self.create_side_by_side_tab()
        self.results_tabs.addTab(side_by_side_tab, "Side by Side")

        # Combined view tab
        combined_tab = self.create_combined_tab()
        self.results_tabs.addTab(combined_tab, "Combined View")

        # Differences list tab
        diff_list_tab = self.create_differences_list_tab()
        self.results_tabs.addTab(diff_list_tab, "Differences")

        layout.addWidget(self.results_tabs)

        # Buttons
        button_layout = QHBoxLayout()

        export_button = QPushButton("Export Report")
        export_button.clicked.connect(self.export_report)
        button_layout.addWidget(export_button)

        button_layout.addStretch()

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def create_summary_tab(self):
        """Create the summary tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setHtml("<p>No comparison performed yet.</p>")

        layout.addWidget(self.summary_text)
        tab.setLayout(layout)
        return tab

    def create_side_by_side_tab(self):
        """Create the side-by-side view tab."""
        tab = QWidget()
        layout = QHBoxLayout()

        # Original document
        original_group = QGroupBox("Original Document")
        original_layout = QVBoxLayout()
        self.original_text = QTextEdit()
        self.original_text.setReadOnly(True)
        original_layout.addWidget(self.original_text)
        original_group.setLayout(original_layout)

        # Revised document
        revised_group = QGroupBox("Revised Document")
        revised_layout = QVBoxLayout()
        self.revised_text = QTextEdit()
        self.revised_text.setReadOnly(True)
        revised_layout.addWidget(self.revised_text)
        revised_group.setLayout(revised_layout)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(original_group)
        splitter.addWidget(revised_group)

        layout.addWidget(splitter)
        tab.setLayout(layout)
        return tab

    def create_combined_tab(self):
        """Create the combined view tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        legend_layout = QHBoxLayout()
        legend_layout.addWidget(QLabel("Legend:"))

        addition_label = QLabel("Addition")
        addition_label.setStyleSheet("background-color: #c8ffc8; color: #008000; padding: 3px;")
        legend_layout.addWidget(addition_label)

        deletion_label = QLabel("Deletion")
        deletion_label.setStyleSheet("background-color: #ffc8c8; color: #ff0000; padding: 3px; text-decoration: line-through;")
        legend_layout.addWidget(deletion_label)

        modification_label = QLabel("Modification")
        modification_label.setStyleSheet("background-color: #ffffc8; color: #808000; padding: 3px;")
        legend_layout.addWidget(modification_label)

        legend_layout.addStretch()
        layout.addLayout(legend_layout)

        self.combined_text = QTextEdit()
        self.combined_text.setReadOnly(True)
        layout.addWidget(self.combined_text)

        tab.setLayout(layout)
        return tab

    def create_differences_list_tab(self):
        """Create the differences list tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        self.differences_list = QListWidget()
        layout.addWidget(self.differences_list)

        tab.setLayout(layout)
        return tab

    def browse_file(self):
        """Browse for a file to compare."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Document to Compare",
            "",
            "Text Files (*.txt);;All Files (*.*)"
        )

        if file_path:
            self.file_path = file_path
            self.file_label.setText(file_path)

    def compare_documents(self):
        """Perform document comparison."""
        if not hasattr(self, 'file_path'):
            QMessageBox.warning(self, "No File", "Please select a document to compare.")
            return

        try:
            # Load the comparison document
            with open(self.file_path, 'r', encoding='utf-8') as f:
                comparison_text = f.read()

            # Create a QTextDocument for comparison
            comparison_doc = QTextDocument()
            comparison_doc.setPlainText(comparison_text)

            # Create comparison object
            self.comparison = DocumentComparison(self.current_document, comparison_doc)

            # Perform comparison
            if self.word_compare_radio.isChecked():
                self.comparison.compare_detailed()
            else:
                self.comparison.compare()

            # Update UI
            self.update_results()

            QMessageBox.information(self, "Success", "Documents compared successfully!")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to compare documents: {str(e)}")

    def update_results(self):
        """Update the results display."""
        if not self.comparison:
            return

        # Update summary
        summary = self.comparison.get_summary()
        summary_html = f"""
        <h3>Comparison Summary</h3>
        <p><b>Total Changes:</b> {summary['total']}</p>
        <p><b>Additions:</b> <span style='color: green;'>{summary['additions']}</span></p>
        <p><b>Deletions:</b> <span style='color: red;'>{summary['deletions']}</span></p>
        <p><b>Modifications:</b> <span style='color: orange;'>{summary['modifications']}</span></p>
        <p><b>Comparison Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        """
        self.summary_text.setHtml(summary_html)

        # Update side-by-side view
        self.original_text.setPlainText(self.comparison.original_doc.toPlainText())
        self.revised_text.setPlainText(self.comparison.revised_doc.toPlainText())

        # Update combined view
        self.combined_text.clear()
        self.comparison.apply_highlighting(
            self.combined_text,
            show_original=self.show_deletions_checkbox.isChecked()
        )

        # Update differences list
        self.differences_list.clear()
        for i, diff in enumerate(self.comparison.differences):
            if diff.type == DifferenceType.ADDITION:
                text = f"[{i+1}] Addition: {diff.new_text[:50]}..."
                item = QListWidgetItem(text)
                item.setForeground(QColor(0, 128, 0))
            elif diff.type == DifferenceType.DELETION:
                text = f"[{i+1}] Deletion: {diff.original_text[:50]}..."
                item = QListWidgetItem(text)
                item.setForeground(QColor(255, 0, 0))
            elif diff.type == DifferenceType.MODIFICATION:
                text = f"[{i+1}] Modification: {diff.original_text[:25]}... → {diff.new_text[:25]}..."
                item = QListWidgetItem(text)
                item.setForeground(QColor(128, 128, 0))
            else:
                continue

            self.differences_list.addItem(item)

    def export_report(self):
        """Export comparison report."""
        if not self.comparison:
            QMessageBox.warning(self, "No Comparison", "Please compare documents first.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Comparison Report",
            "",
            "HTML Files (*.html);;Text Files (*.txt)"
        )

        if file_path:
            try:
                summary = self.comparison.get_summary()

                if file_path.endswith('.html'):
                    # Export as HTML
                    html_content = f"""
                    <html>
                    <head>
                        <title>Document Comparison Report</title>
                        <style>
                            body {{ font-family: Arial, sans-serif; margin: 20px; }}
                            h1 {{ color: #333; }}
                            .addition {{ background-color: #c8ffc8; color: #008000; }}
                            .deletion {{ background-color: #ffc8c8; color: #ff0000; text-decoration: line-through; }}
                            .modification {{ background-color: #ffffc8; color: #808000; }}
                            .summary {{ margin: 20px 0; padding: 10px; background-color: #f0f0f0; }}
                        </style>
                    </head>
                    <body>
                        <h1>Document Comparison Report</h1>
                        <div class="summary">
                            <h2>Summary</h2>
                            <p><b>Total Changes:</b> {summary['total']}</p>
                            <p><b>Additions:</b> {summary['additions']}</p>
                            <p><b>Deletions:</b> {summary['deletions']}</p>
                            <p><b>Modifications:</b> {summary['modifications']}</p>
                            <p><b>Report Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                        </div>
                        <h2>Differences</h2>
                        <div>
                    """

                    for i, diff in enumerate(self.comparison.differences):
                        if diff.type == DifferenceType.ADDITION:
                            html_content += f'<p><span class="addition">[Addition {i+1}]: {diff.new_text}</span></p>'
                        elif diff.type == DifferenceType.DELETION:
                            html_content += f'<p><span class="deletion">[Deletion {i+1}]: {diff.original_text}</span></p>'
                        elif diff.type == DifferenceType.MODIFICATION:
                            html_content += f'<p><span class="modification">[Modification {i+1}]: {diff.original_text} → {diff.new_text}</span></p>'

                    html_content += """
                        </div>
                    </body>
                    </html>
                    """

                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(html_content)

                else:
                    # Export as plain text
                    text_content = "Document Comparison Report\n"
                    text_content += "=" * 50 + "\n\n"
                    text_content += f"Total Changes: {summary['total']}\n"
                    text_content += f"Additions: {summary['additions']}\n"
                    text_content += f"Deletions: {summary['deletions']}\n"
                    text_content += f"Modifications: {summary['modifications']}\n"
                    text_content += f"Report Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
                    text_content += "Differences:\n"
                    text_content += "-" * 50 + "\n\n"

                    for i, diff in enumerate(self.comparison.differences):
                        if diff.type == DifferenceType.ADDITION:
                            text_content += f"[Addition {i+1}]: {diff.new_text}\n\n"
                        elif diff.type == DifferenceType.DELETION:
                            text_content += f"[Deletion {i+1}]: {diff.original_text}\n\n"
                        elif diff.type == DifferenceType.MODIFICATION:
                            text_content += f"[Modification {i+1}]: {diff.original_text} → {diff.new_text}\n\n"

                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(text_content)

                QMessageBox.information(self, "Success", "Report exported successfully!")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to export report: {str(e)}")
