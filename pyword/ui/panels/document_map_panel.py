"""Document Map Panel for PyWord."""

from PySide6.QtWidgets import (QWidget, QVBoxLayout, QTreeWidget, QTreeWidgetItem,
                             QLabel, QSizePolicy)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont


class DocumentMapPanel(QWidget):
    """
    Document map panel that shows an outline view of the document structure.

    Displays headings and allows quick navigation to different sections.
    """

    # Signal emitted when a heading is clicked
    heading_clicked = Signal(int)  # Emits the line number

    def __init__(self, parent=None):
        super().__init__(parent)
        self.document = None
        self.setup_ui()

    def setup_ui(self):
        """Initialize the panel UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)

        # Title
        title_label = QLabel("Document Map")
        title_font = QFont()
        title_font.setBold(True)
        title_label.setFont(title_font)
        layout.addWidget(title_label)

        # Tree widget for headings
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setIndentation(15)
        self.tree.itemClicked.connect(self.on_item_clicked)
        layout.addWidget(self.tree)

        # Info label
        self.info_label = QLabel("No headings in document")
        self.info_label.setAlignment(Qt.AlignCenter)
        self.info_label.setStyleSheet("color: gray;")
        layout.addWidget(self.info_label)

        # Initially show info, hide tree
        self.tree.hide()

    def set_document(self, document):
        """Set the document to display in the map."""
        self.document = document
        self.update_map()

    def update_map(self):
        """Update the document map based on current document content."""
        self.tree.clear()

        if not self.document:
            self.tree.hide()
            self.info_label.show()
            self.info_label.setText("No document loaded")
            return

        # Parse document for headings
        headings = self.extract_headings()

        if not headings:
            self.tree.hide()
            self.info_label.show()
            self.info_label.setText("No headings in document")
            return

        # Populate tree
        self.info_label.hide()
        self.tree.show()

        # Stack to keep track of parent items for each heading level
        parent_stack = [None]  # Root level
        last_items = [None] * 10  # Support up to 10 heading levels

        for heading in headings:
            level = heading['level']
            text = heading['text']
            line_number = heading['line']

            # Create tree item
            item = QTreeWidgetItem([text])
            item.setData(0, Qt.UserRole, line_number)

            # Set font based on heading level
            font = QFont()
            if level == 1:
                font.setBold(True)
                font.setPointSize(12)
            elif level == 2:
                font.setBold(True)
                font.setPointSize(11)
            else:
                font.setPointSize(10)
            item.setFont(0, font)

            # Find appropriate parent
            if level == 1:
                self.tree.addTopLevelItem(item)
                parent_stack = [None]
                last_items[0] = item
            else:
                # Find the parent item (last item at level-1)
                parent = last_items[level - 2] if level >= 2 else None
                if parent:
                    parent.addChild(item)
                else:
                    self.tree.addTopLevelItem(item)
                last_items[level - 1] = item

        # Expand all items
        self.tree.expandAll()

    def extract_headings(self):
        """
        Extract headings from the document.

        Returns:
            List of dictionaries with 'level', 'text', and 'line' keys
        """
        if not self.document:
            return []

        headings = []

        # Get document content
        # This is a simplified version - in a real implementation,
        # you'd parse the actual document format
        try:
            if hasattr(self.document, 'toPlainText'):
                text = self.document.toPlainText()
            elif hasattr(self.document, 'content'):
                text = self.document.content
            else:
                return []

            lines = text.split('\n')

            for line_num, line in enumerate(lines, 1):
                line = line.strip()

                # Detect markdown-style headings
                if line.startswith('#'):
                    level = len(line) - len(line.lstrip('#'))
                    text = line.lstrip('#').strip()
                    if text:
                        headings.append({
                            'level': min(level, 6),
                            'text': text,
                            'line': line_num
                        })

                # Detect all-caps lines as potential headings
                elif line.isupper() and len(line) > 3 and len(line) < 100:
                    headings.append({
                        'level': 1,
                        'text': line,
                        'line': line_num
                    })

        except Exception as e:
            print(f"Error extracting headings: {e}")

        return headings

    def on_item_clicked(self, item, column):
        """Handle heading item click."""
        line_number = item.data(0, Qt.UserRole)
        if line_number:
            self.heading_clicked.emit(line_number)

    def refresh(self):
        """Refresh the document map."""
        self.update_map()
