"""
Citations and Bibliography System for PyWord.

This module provides comprehensive support for citations and bibliography management,
including multiple citation styles (APA, MLA, Chicago, etc.) and automatic formatting.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QListWidget, QListWidgetItem, QGroupBox,
                               QComboBox, QTextEdit, QFormLayout, QMessageBox,
                               QTabWidget, QWidget, QCheckBox, QSpinBox, QTableWidget,
                               QTableWidgetItem, QHeaderView, QFileDialog)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QTextCursor, QTextCharFormat, QFont, QColor
from datetime import datetime
import json
import uuid


class Citation:
    """Represents a bibliographic citation."""

    def __init__(self, citation_type='book'):
        self.id = str(uuid.uuid4())
        self.citation_type = citation_type  # book, article, website, etc.

        # Common fields
        self.author = ""
        self.title = ""
        self.year = ""

        # Book-specific
        self.publisher = ""
        self.place = ""
        self.edition = ""
        self.isbn = ""

        # Article-specific
        self.journal = ""
        self.volume = ""
        self.issue = ""
        self.pages = ""
        self.doi = ""

        # Website-specific
        self.url = ""
        self.access_date = ""

        # Additional fields
        self.editor = ""
        self.translator = ""
        self.notes = ""

        self.created = datetime.now()
        self.used_count = 0  # Number of times cited in document

    def to_dict(self):
        """Convert citation to dictionary for serialization."""
        return {
            'id': self.id,
            'citation_type': self.citation_type,
            'author': self.author,
            'title': self.title,
            'year': self.year,
            'publisher': self.publisher,
            'place': self.place,
            'edition': self.edition,
            'isbn': self.isbn,
            'journal': self.journal,
            'volume': self.volume,
            'issue': self.issue,
            'pages': self.pages,
            'doi': self.doi,
            'url': self.url,
            'access_date': self.access_date,
            'editor': self.editor,
            'translator': self.translator,
            'notes': self.notes,
            'created': self.created.isoformat(),
            'used_count': self.used_count
        }

    @staticmethod
    def from_dict(data):
        """Create Citation from dictionary."""
        citation = Citation(data.get('citation_type', 'book'))
        citation.id = data['id']
        citation.author = data.get('author', '')
        citation.title = data.get('title', '')
        citation.year = data.get('year', '')
        citation.publisher = data.get('publisher', '')
        citation.place = data.get('place', '')
        citation.edition = data.get('edition', '')
        citation.isbn = data.get('isbn', '')
        citation.journal = data.get('journal', '')
        citation.volume = data.get('volume', '')
        citation.issue = data.get('issue', '')
        citation.pages = data.get('pages', '')
        citation.doi = data.get('doi', '')
        citation.url = data.get('url', '')
        citation.access_date = data.get('access_date', '')
        citation.editor = data.get('editor', '')
        citation.translator = data.get('translator', '')
        citation.notes = data.get('notes', '')
        if 'created' in data:
            citation.created = datetime.fromisoformat(data['created'])
        citation.used_count = data.get('used_count', 0)
        return citation


class CitationReference:
    """Represents a reference to a citation in the document."""

    def __init__(self, citation_id, position, page_number=1):
        self.id = str(uuid.uuid4())
        self.citation_id = citation_id
        self.position = position
        self.page_number = page_number


class CitationsManager:
    """Manages citations and bibliography in a document."""

    def __init__(self, parent):
        self.parent = parent  # Reference to text editor
        self.citations = []  # All available citations
        self.references = []  # References to citations in document

        # Settings
        self.citation_style = 'APA'  # APA, MLA, Chicago, Harvard, IEEE, etc.
        self.sort_bibliography = True
        self.sort_by = 'author'  # author, title, year
        self.indent_hanging = True

    def add_citation(self, citation):
        """Add a new citation to the library."""
        self.citations.append(citation)
        return citation

    def edit_citation(self, citation_id, updated_citation):
        """Update an existing citation."""
        citation = self.get_citation_by_id(citation_id)
        if citation:
            # Update fields
            for key, value in updated_citation.to_dict().items():
                if hasattr(citation, key) and key != 'id':
                    setattr(citation, key, value)
            return True
        return False

    def delete_citation(self, citation_id):
        """Delete a citation from the library."""
        citation = self.get_citation_by_id(citation_id)
        if citation:
            # Remove all references to this citation
            self.references = [ref for ref in self.references if ref.citation_id != citation_id]
            self.citations.remove(citation)
            return True
        return False

    def get_citation_by_id(self, citation_id):
        """Get a citation by its ID."""
        for citation in self.citations:
            if citation.id == citation_id:
                return citation
        return None

    def insert_citation(self, citation_id, position=None):
        """Insert a citation reference in the document."""
        if position is None:
            cursor = self.parent.textCursor()
            position = cursor.position()

        citation = self.get_citation_by_id(citation_id)
        if not citation:
            return None

        # Create reference
        page_number = 1  # Simplified - calculate actual page
        reference = CitationReference(citation_id, position, page_number)
        self.references.append(reference)

        # Increment usage count
        citation.used_count += 1

        # Insert formatted citation in document
        self._insert_citation_mark(citation, position)

        return reference

    def _insert_citation_mark(self, citation, position):
        """Insert citation mark in the document."""
        cursor = QTextCursor(self.parent.document())
        cursor.setPosition(position)

        # Format citation according to style
        citation_text = self._format_inline_citation(citation)

        # Format for citation mark
        char_format = QTextCharFormat()
        char_format.setForeground(QColor(0, 100, 0))  # Dark green

        cursor.setCharFormat(char_format)
        cursor.insertText(citation_text)

    def _format_inline_citation(self, citation):
        """Format an inline citation according to the selected style."""
        if self.citation_style == 'APA':
            return self._format_apa_inline(citation)
        elif self.citation_style == 'MLA':
            return self._format_mla_inline(citation)
        elif self.citation_style == 'Chicago':
            return self._format_chicago_inline(citation)
        elif self.citation_style == 'Harvard':
            return self._format_harvard_inline(citation)
        elif self.citation_style == 'IEEE':
            return self._format_ieee_inline(citation)
        else:
            return f" ({citation.author}, {citation.year})"

    def _format_apa_inline(self, citation):
        """Format APA style inline citation."""
        author_part = citation.author.split(',')[0] if citation.author else "Unknown"
        return f" ({author_part}, {citation.year})"

    def _format_mla_inline(self, citation):
        """Format MLA style inline citation."""
        author_part = citation.author.split(',')[0] if citation.author else "Unknown"
        if citation.pages:
            return f" ({author_part} {citation.pages})"
        return f" ({author_part})"

    def _format_chicago_inline(self, citation):
        """Format Chicago style inline citation."""
        author_part = citation.author.split(',')[0] if citation.author else "Unknown"
        return f" ({author_part} {citation.year})"

    def _format_harvard_inline(self, citation):
        """Format Harvard style inline citation."""
        author_part = citation.author.split(',')[0] if citation.author else "Unknown"
        return f" ({author_part} {citation.year})"

    def _format_ieee_inline(self, citation):
        """Format IEEE style inline citation (numbered)."""
        # Find citation number in bibliography
        sorted_citations = self._get_sorted_citations()
        try:
            number = sorted_citations.index(citation) + 1
            return f" [{number}]"
        except ValueError:
            return " [?]"

    def insert_bibliography(self, position=None):
        """Insert a bibliography/references section."""
        if position is None:
            cursor = self.parent.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.End)
        else:
            cursor = QTextCursor(self.parent.document())
            cursor.setPosition(position)

        # Get citations used in document
        used_citations = self._get_used_citations()

        if not used_citations:
            QMessageBox.warning(
                None,
                "No Citations",
                "No citations found in the document. Please insert citations before generating a bibliography."
            )
            return False

        # Insert section title
        cursor.insertBlock()
        title_format = QTextCharFormat()
        title_format.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        cursor.setCharFormat(title_format)

        # Title varies by style
        if self.citation_style in ['APA', 'Harvard']:
            title = "References"
        elif self.citation_style == 'MLA':
            title = "Works Cited"
        elif self.citation_style == 'Chicago':
            title = "Bibliography"
        else:
            title = "References"

        cursor.insertText(title)
        cursor.insertBlock()

        # Insert citations
        sorted_citations = self._get_sorted_citations(used_citations)

        for i, citation in enumerate(sorted_citations):
            formatted = self._format_bibliography_entry(citation, i + 1)
            self._insert_bibliography_entry(cursor, formatted)

        return True

    def _get_used_citations(self):
        """Get list of citations that are referenced in the document."""
        used_ids = set(ref.citation_id for ref in self.references)
        return [c for c in self.citations if c.id in used_ids]

    def _get_sorted_citations(self, citations=None):
        """Get sorted list of citations based on settings."""
        if citations is None:
            citations = self.citations

        if not self.sort_bibliography:
            return citations

        if self.sort_by == 'author':
            return sorted(citations, key=lambda c: c.author.lower() if c.author else '')
        elif self.sort_by == 'title':
            return sorted(citations, key=lambda c: c.title.lower() if c.title else '')
        elif self.sort_by == 'year':
            return sorted(citations, key=lambda c: c.year if c.year else '')

        return citations

    def _format_bibliography_entry(self, citation, number=None):
        """Format a bibliography entry according to the selected style."""
        if self.citation_style == 'APA':
            return self._format_apa_bibliography(citation)
        elif self.citation_style == 'MLA':
            return self._format_mla_bibliography(citation)
        elif self.citation_style == 'Chicago':
            return self._format_chicago_bibliography(citation)
        elif self.citation_style == 'Harvard':
            return self._format_harvard_bibliography(citation)
        elif self.citation_style == 'IEEE':
            return self._format_ieee_bibliography(citation, number)
        else:
            return f"{citation.author}. ({citation.year}). {citation.title}."

    def _format_apa_bibliography(self, citation):
        """Format APA style bibliography entry."""
        parts = []

        if citation.author:
            parts.append(f"{citation.author}.")

        if citation.year:
            parts.append(f"({citation.year}).")

        if citation.title:
            if citation.citation_type == 'book':
                parts.append(f"<i>{citation.title}</i>.")
            else:
                parts.append(f"{citation.title}.")

        if citation.citation_type == 'book':
            if citation.publisher:
                parts.append(f"{citation.publisher}.")
        elif citation.citation_type == 'article':
            if citation.journal:
                journal_part = f"<i>{citation.journal}</i>"
                if citation.volume:
                    journal_part += f", {citation.volume}"
                if citation.issue:
                    journal_part += f"({citation.issue})"
                if citation.pages:
                    journal_part += f", {citation.pages}"
                parts.append(journal_part + ".")

        if citation.doi:
            parts.append(f"https://doi.org/{citation.doi}")

        return " ".join(parts)

    def _format_mla_bibliography(self, citation):
        """Format MLA style bibliography entry."""
        parts = []

        if citation.author:
            parts.append(f"{citation.author}.")

        if citation.title:
            if citation.citation_type == 'book':
                parts.append(f"<i>{citation.title}</i>.")
            else:
                parts.append(f'"{citation.title}."')

        if citation.citation_type == 'article' and citation.journal:
            parts.append(f"<i>{citation.journal}</i>,")
            if citation.volume:
                parts.append(f"vol. {citation.volume},")
            if citation.issue:
                parts.append(f"no. {citation.issue},")

        if citation.publisher:
            parts.append(f"{citation.publisher},")

        if citation.year:
            parts.append(f"{citation.year}.")

        if citation.pages:
            parts.append(f"pp. {citation.pages}.")

        return " ".join(parts)

    def _format_chicago_bibliography(self, citation):
        """Format Chicago style bibliography entry."""
        parts = []

        if citation.author:
            parts.append(f"{citation.author}.")

        if citation.title:
            if citation.citation_type == 'book':
                parts.append(f"<i>{citation.title}</i>.")
            else:
                parts.append(f'"{citation.title}."')

        if citation.citation_type == 'article' and citation.journal:
            journal_part = f"<i>{citation.journal}</i>"
            if citation.volume:
                journal_part += f" {citation.volume}"
            if citation.issue:
                journal_part += f", no. {citation.issue}"
            parts.append(journal_part)

        if citation.place:
            parts.append(f"{citation.place}:")

        if citation.publisher:
            parts.append(f"{citation.publisher},")

        if citation.year:
            parts.append(f"{citation.year}.")

        return " ".join(parts)

    def _format_harvard_bibliography(self, citation):
        """Format Harvard style bibliography entry."""
        return self._format_apa_bibliography(citation)  # Similar to APA

    def _format_ieee_bibliography(self, citation, number):
        """Format IEEE style bibliography entry."""
        parts = [f"[{number}]"]

        if citation.author:
            # IEEE uses initials first
            parts.append(f"{citation.author},")

        if citation.title:
            parts.append(f'"{citation.title},"')

        if citation.citation_type == 'article' and citation.journal:
            journal_part = f"<i>{citation.journal}</i>"
            if citation.volume:
                journal_part += f", vol. {citation.volume}"
            if citation.issue:
                journal_part += f", no. {citation.issue}"
            if citation.pages:
                journal_part += f", pp. {citation.pages}"
            parts.append(journal_part + ",")

        if citation.year:
            parts.append(f"{citation.year}.")

        return " ".join(parts)

    def _insert_bibliography_entry(self, cursor, formatted_text):
        """Insert a formatted bibliography entry."""
        # Create hanging indent if enabled
        if self.indent_hanging:
            block_format = cursor.blockFormat()
            block_format.setTextIndent(-20)
            block_format.setLeftMargin(20)
            cursor.setBlockFormat(block_format)

        cursor.insertBlock()

        # Insert formatted text (handle simple HTML tags)
        self._insert_formatted_text(cursor, formatted_text)

    def _insert_formatted_text(self, cursor, text):
        """Insert text with simple HTML formatting."""
        # Simple HTML parsing for italic tags
        import re

        parts = re.split(r'(<i>.*?</i>)', text)

        for part in parts:
            if part.startswith('<i>') and part.endswith('</i>'):
                # Italic text
                char_format = QTextCharFormat()
                char_format.setFontItalic(True)
                cursor.setCharFormat(char_format)
                cursor.insertText(part[3:-4])  # Remove <i> and </i>

                # Reset format
                char_format.setFontItalic(False)
                cursor.setCharFormat(char_format)
            else:
                cursor.insertText(part)

    def export_citations(self, file_path, format='bibtex'):
        """Export citations to various formats."""
        if format == 'bibtex':
            return self._export_bibtex(file_path)
        elif format == 'json':
            return self._export_json(file_path)
        return False

    def _export_bibtex(self, file_path):
        """Export citations to BibTeX format."""
        try:
            with open(file_path, 'w') as f:
                for citation in self.citations:
                    entry_type = citation.citation_type
                    cite_key = f"{citation.author.split(',')[0]}{citation.year}".replace(' ', '')

                    f.write(f"@{entry_type}{{{cite_key},\n")
                    if citation.author:
                        f.write(f"  author = {{{citation.author}}},\n")
                    if citation.title:
                        f.write(f"  title = {{{citation.title}}},\n")
                    if citation.year:
                        f.write(f"  year = {{{citation.year}}},\n")
                    if citation.publisher:
                        f.write(f"  publisher = {{{citation.publisher}}},\n")
                    if citation.journal:
                        f.write(f"  journal = {{{citation.journal}}},\n")
                    if citation.volume:
                        f.write(f"  volume = {{{citation.volume}}},\n")
                    if citation.pages:
                        f.write(f"  pages = {{{citation.pages}}},\n")
                    f.write("}\n\n")
            return True
        except Exception as e:
            print(f"Error exporting to BibTeX: {e}")
            return False

    def _export_json(self, file_path):
        """Export citations to JSON format."""
        try:
            data = {
                'citations': [c.to_dict() for c in self.citations],
                'style': self.citation_style,
                'exported': datetime.now().isoformat()
            }
            with open(file_path, 'w') as f:
                json.dump(data, f, indent=2)
            return True
        except Exception as e:
            print(f"Error exporting to JSON: {e}")
            return False

    def import_citations(self, file_path):
        """Import citations from JSON format."""
        try:
            with open(file_path, 'r') as f:
                data = json.load(f)

            for citation_data in data.get('citations', []):
                citation = Citation.from_dict(citation_data)
                self.citations.append(citation)

            return True
        except Exception as e:
            print(f"Error importing citations: {e}")
            return False


class CitationDialog(QDialog):
    """Dialog for adding/editing a citation."""

    def __init__(self, citation=None, parent=None):
        super().__init__(parent)
        self.citation = citation if citation else Citation()

        title = "Edit Citation" if citation else "Add Citation"
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumSize(600, 500)

        self.setup_ui()
        self.load_citation()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Citation type
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Type:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(['book', 'article', 'website', 'conference', 'thesis', 'other'])
        self.type_combo.currentTextChanged.connect(self.on_type_changed)
        type_layout.addWidget(self.type_combo)
        type_layout.addStretch()
        layout.addWidget(QLabel("<h3>Citation Details</h3>"))
        layout.addLayout(type_layout)

        # Tabs for different field groups
        self.tabs = QTabWidget()

        # Basic tab
        basic_tab = QWidget()
        basic_layout = QFormLayout()

        self.author_edit = QLineEdit()
        self.author_edit.setPlaceholderText("Last, First; Last, First")
        basic_layout.addRow("Author(s):", self.author_edit)

        self.title_edit = QLineEdit()
        basic_layout.addRow("Title:", self.title_edit)

        self.year_edit = QLineEdit()
        self.year_edit.setPlaceholderText("YYYY")
        basic_layout.addRow("Year:", self.year_edit)

        basic_tab.setLayout(basic_layout)
        self.tabs.addTab(basic_tab, "Basic")

        # Publication tab
        pub_tab = QWidget()
        pub_layout = QFormLayout()

        self.publisher_edit = QLineEdit()
        pub_layout.addRow("Publisher:", self.publisher_edit)

        self.place_edit = QLineEdit()
        pub_layout.addRow("Place:", self.place_edit)

        self.edition_edit = QLineEdit()
        pub_layout.addRow("Edition:", self.edition_edit)

        self.isbn_edit = QLineEdit()
        pub_layout.addRow("ISBN:", self.isbn_edit)

        pub_tab.setLayout(pub_layout)
        self.tabs.addTab(pub_tab, "Publication")

        # Journal tab
        journal_tab = QWidget()
        journal_layout = QFormLayout()

        self.journal_edit = QLineEdit()
        journal_layout.addRow("Journal:", self.journal_edit)

        self.volume_edit = QLineEdit()
        journal_layout.addRow("Volume:", self.volume_edit)

        self.issue_edit = QLineEdit()
        journal_layout.addRow("Issue:", self.issue_edit)

        self.pages_edit = QLineEdit()
        self.pages_edit.setPlaceholderText("e.g., 123-456")
        journal_layout.addRow("Pages:", self.pages_edit)

        self.doi_edit = QLineEdit()
        journal_layout.addRow("DOI:", self.doi_edit)

        journal_tab.setLayout(journal_layout)
        self.tabs.addTab(journal_tab, "Journal")

        # Online tab
        online_tab = QWidget()
        online_layout = QFormLayout()

        self.url_edit = QLineEdit()
        online_layout.addRow("URL:", self.url_edit)

        self.access_date_edit = QLineEdit()
        self.access_date_edit.setPlaceholderText("YYYY-MM-DD")
        online_layout.addRow("Access Date:", self.access_date_edit)

        online_tab.setLayout(online_layout)
        self.tabs.addTab(online_tab, "Online")

        # Other tab
        other_tab = QWidget()
        other_layout = QFormLayout()

        self.editor_edit = QLineEdit()
        other_layout.addRow("Editor:", self.editor_edit)

        self.translator_edit = QLineEdit()
        other_layout.addRow("Translator:", self.translator_edit)

        self.notes_edit = QTextEdit()
        self.notes_edit.setMaximumHeight(80)
        other_layout.addRow("Notes:", self.notes_edit)

        other_tab.setLayout(other_layout)
        self.tabs.addTab(other_tab, "Other")

        layout.addWidget(self.tabs)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.save_citation)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def on_type_changed(self):
        """Handle citation type change."""
        # Could enable/disable relevant tabs based on type
        pass

    def load_citation(self):
        """Load citation data into form."""
        self.type_combo.setCurrentText(self.citation.citation_type)
        self.author_edit.setText(self.citation.author)
        self.title_edit.setText(self.citation.title)
        self.year_edit.setText(self.citation.year)
        self.publisher_edit.setText(self.citation.publisher)
        self.place_edit.setText(self.citation.place)
        self.edition_edit.setText(self.citation.edition)
        self.isbn_edit.setText(self.citation.isbn)
        self.journal_edit.setText(self.citation.journal)
        self.volume_edit.setText(self.citation.volume)
        self.issue_edit.setText(self.citation.issue)
        self.pages_edit.setText(self.citation.pages)
        self.doi_edit.setText(self.citation.doi)
        self.url_edit.setText(self.citation.url)
        self.access_date_edit.setText(self.citation.access_date)
        self.editor_edit.setText(self.citation.editor)
        self.translator_edit.setText(self.citation.translator)
        self.notes_edit.setPlainText(self.citation.notes)

    def save_citation(self):
        """Save form data to citation."""
        if not self.author_edit.text() or not self.title_edit.text():
            QMessageBox.warning(self, "Required Fields", "Author and Title are required.")
            return

        self.citation.citation_type = self.type_combo.currentText()
        self.citation.author = self.author_edit.text()
        self.citation.title = self.title_edit.text()
        self.citation.year = self.year_edit.text()
        self.citation.publisher = self.publisher_edit.text()
        self.citation.place = self.place_edit.text()
        self.citation.edition = self.edition_edit.text()
        self.citation.isbn = self.isbn_edit.text()
        self.citation.journal = self.journal_edit.text()
        self.citation.volume = self.volume_edit.text()
        self.citation.issue = self.issue_edit.text()
        self.citation.pages = self.pages_edit.text()
        self.citation.doi = self.doi_edit.text()
        self.citation.url = self.url_edit.text()
        self.citation.access_date = self.access_date_edit.text()
        self.citation.editor = self.editor_edit.text()
        self.citation.translator = self.translator_edit.text()
        self.citation.notes = self.notes_edit.toPlainText()

        self.accept()

    def get_citation(self):
        """Get the citation object."""
        return self.citation


class CitationsManagerDialog(QDialog):
    """Dialog for managing citations library and inserting citations."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Citations and Bibliography")
        self.setModal(False)
        self.setMinimumSize(800, 600)

        self.setup_ui()
        self.refresh_citations()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Header
        header_label = QLabel("<h2>Citations and Bibliography</h2>")
        layout.addWidget(header_label)

        # Style selection
        style_layout = QHBoxLayout()
        style_layout.addWidget(QLabel("Citation Style:"))
        self.style_combo = QComboBox()
        self.style_combo.addItems(['APA', 'MLA', 'Chicago', 'Harvard', 'IEEE'])
        self.style_combo.setCurrentText(self.manager.citation_style)
        self.style_combo.currentTextChanged.connect(self.on_style_changed)
        style_layout.addWidget(self.style_combo)
        style_layout.addStretch()
        layout.addLayout(style_layout)

        # Citations list
        list_label = QLabel("<b>Citation Library:</b>")
        layout.addWidget(list_label)

        self.citations_table = QTableWidget()
        self.citations_table.setColumnCount(4)
        self.citations_table.setHorizontalHeaderLabels(['Author', 'Title', 'Year', 'Type'])
        self.citations_table.horizontalHeader().setStretchLastSection(False)
        self.citations_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.citations_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        layout.addWidget(self.citations_table)

        # Action buttons
        button_layout = QHBoxLayout()

        add_button = QPushButton("Add Citation")
        add_button.clicked.connect(self.add_citation)
        button_layout.addWidget(add_button)

        edit_button = QPushButton("Edit")
        edit_button.clicked.connect(self.edit_citation)
        button_layout.addWidget(edit_button)

        delete_button = QPushButton("Delete")
        delete_button.clicked.connect(self.delete_citation)
        button_layout.addWidget(delete_button)

        button_layout.addStretch()

        insert_button = QPushButton("Insert Citation")
        insert_button.clicked.connect(self.insert_citation)
        button_layout.addWidget(insert_button)

        bibliography_button = QPushButton("Insert Bibliography")
        bibliography_button.clicked.connect(self.insert_bibliography)
        button_layout.addWidget(bibliography_button)

        layout.addLayout(button_layout)

        # Import/Export buttons
        ie_layout = QHBoxLayout()

        import_button = QPushButton("Import")
        import_button.clicked.connect(self.import_citations)
        ie_layout.addWidget(import_button)

        export_button = QPushButton("Export")
        export_button.clicked.connect(self.export_citations)
        ie_layout.addWidget(export_button)

        ie_layout.addStretch()

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        ie_layout.addWidget(close_button)

        layout.addLayout(ie_layout)

        self.setLayout(layout)

    def on_style_changed(self, style):
        """Handle citation style change."""
        self.manager.citation_style = style

    def refresh_citations(self):
        """Refresh the citations table."""
        self.citations_table.setRowCount(0)

        for citation in self.manager.citations:
            row = self.citations_table.rowCount()
            self.citations_table.insertRow(row)

            self.citations_table.setItem(row, 0, QTableWidgetItem(citation.author))
            self.citations_table.setItem(row, 1, QTableWidgetItem(citation.title))
            self.citations_table.setItem(row, 2, QTableWidgetItem(citation.year))
            self.citations_table.setItem(row, 3, QTableWidgetItem(citation.citation_type))

            # Store citation ID in row
            self.citations_table.item(row, 0).setData(Qt.ItemDataRole.UserRole, citation.id)

    def add_citation(self):
        """Add a new citation."""
        dialog = CitationDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            citation = dialog.get_citation()
            self.manager.add_citation(citation)
            self.refresh_citations()

    def edit_citation(self):
        """Edit the selected citation."""
        row = self.citations_table.currentRow()
        if row < 0:
            return

        citation_id = self.citations_table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        citation = self.manager.get_citation_by_id(citation_id)

        if citation:
            dialog = CitationDialog(citation, parent=self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.refresh_citations()

    def delete_citation(self):
        """Delete the selected citation."""
        row = self.citations_table.currentRow()
        if row < 0:
            return

        reply = QMessageBox.question(
            self,
            "Delete Citation",
            "Are you sure you want to delete this citation?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            citation_id = self.citations_table.item(row, 0).data(Qt.ItemDataRole.UserRole)
            self.manager.delete_citation(citation_id)
            self.refresh_citations()

    def insert_citation(self):
        """Insert the selected citation into the document."""
        row = self.citations_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "No Selection", "Please select a citation to insert.")
            return

        citation_id = self.citations_table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        self.manager.insert_citation(citation_id)
        QMessageBox.information(self, "Success", "Citation inserted successfully!")

    def insert_bibliography(self):
        """Insert bibliography into the document."""
        if self.manager.insert_bibliography():
            QMessageBox.information(self, "Success", "Bibliography inserted successfully!")

    def import_citations(self):
        """Import citations from file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Import Citations",
            "",
            "JSON Files (*.json)"
        )

        if file_path:
            if self.manager.import_citations(file_path):
                QMessageBox.information(self, "Success", "Citations imported successfully!")
                self.refresh_citations()
            else:
                QMessageBox.warning(self, "Error", "Failed to import citations.")

    def export_citations(self):
        """Export citations to file."""
        file_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Export Citations",
            "",
            "BibTeX Files (*.bib);;JSON Files (*.json)"
        )

        if file_path:
            format = 'bibtex' if 'BibTeX' in selected_filter else 'json'
            if self.manager.export_citations(file_path, format):
                QMessageBox.information(self, "Success", "Citations exported successfully!")
            else:
                QMessageBox.warning(self, "Error", "Failed to export citations.")
