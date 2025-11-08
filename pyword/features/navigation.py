from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                             QLineEdit, QPushButton, QListWidget, QSplitter,
                             QWidget, QCheckBox)
from PySide6.QtCore import Qt, Signal, QRegularExpression
from PySide6.QtGui import QTextCursor, QTextDocument, QTextBlock, QTextFormat

class FindReplaceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("Find and Replace")
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Find section
        find_layout = QHBoxLayout()
        find_layout.addWidget(QLabel("Find:"))
        self.find_edit = QLineEdit()
        find_layout.addWidget(self.find_edit)
        self.find_prev_btn = QPushButton("Previous")
        self.find_next_btn = QPushButton("Next")
        find_layout.addWidget(self.find_prev_btn)
        find_layout.addWidget(self.find_next_btn)
        
        # Replace section
        replace_layout = QHBoxLayout()
        replace_layout.addWidget(QLabel("Replace with:"))
        self.replace_edit = QLineEdit()
        replace_layout.addWidget(self.replace_edit)
        self.replace_btn = QPushButton("Replace")
        self.replace_all_btn = QPushButton("Replace All")
        replace_layout.addWidget(self.replace_btn)
        replace_layout.addWidget(self.replace_all_btn)
        
        # Options
        options_layout = QHBoxLayout()
        self.match_case = QCheckBox("Match case")
        self.whole_words = QCheckBox("Whole words only")
        self.regex = QCheckBox("Regular expression")
        options_layout.addWidget(self.match_case)
        options_layout.addWidget(self.whole_words)
        options_layout.addWidget(self.regex)

        # Status label
        self.status_label = QLabel("")

        layout.addLayout(find_layout)
        layout.addLayout(replace_layout)
        layout.addLayout(options_layout)
        layout.addWidget(self.status_label)

        self.setLayout(layout)
        
        # Connect signals
        self.find_next_btn.clicked.connect(self.find_next)
        self.find_prev_btn.clicked.connect(self.find_previous)
        self.replace_btn.clicked.connect(self.replace)
        self.replace_all_btn.clicked.connect(self.replace_all)
    
    def find_next(self):
        text = self.find_edit.text()
        if not text:
            return

        if not self.parent or not hasattr(self.parent, 'find'):
            self.status_label.setText("Error: Editor not available")
            return

        flags = QTextDocument.FindFlag(0)
        if self.match_case.isChecked():
            flags |= QTextDocument.FindCaseSensitively
        if self.whole_words.isChecked():
            flags |= QTextDocument.FindWholeWords

        if self.parent.find(text, flags):
            self.status_label.setText("")
        else:
            self.status_label.setText("Reached end of document")
    
    def find_previous(self):
        text = self.find_edit.text()
        if not text:
            return

        if not self.parent or not hasattr(self.parent, 'find'):
            self.status_label.setText("Error: Editor not available")
            return

        flags = QTextDocument.FindFlag(QTextDocument.FindBackward)
        if self.match_case.isChecked():
            flags |= QTextDocument.FindCaseSensitively
        if self.whole_words.isChecked():
            flags |= QTextDocument.FindWholeWords

        if self.parent.find(text, flags):
            self.status_label.setText("")
        else:
            self.status_label.setText("Reached start of document")
    
    def replace(self):
        if not self.parent or not hasattr(self.parent, 'textCursor'):
            self.status_label.setText("Error: Editor not available")
            return

        cursor = self.parent.textCursor()
        if cursor.hasSelection():
            # Compare with case sensitivity based on checkbox
            selected = cursor.selectedText()
            find_text = self.find_edit.text()
            if self.match_case.isChecked():
                match = selected == find_text
            else:
                match = selected.lower() == find_text.lower()

            if match:
                cursor.insertText(self.replace_edit.text())
        self.find_next()
    
    def replace_all(self):
        if not self.parent or not hasattr(self.parent, 'moveCursor') or not hasattr(self.parent, 'find') or not hasattr(self.parent, 'textCursor'):
            self.status_label.setText("Error: Editor not available")
            return

        self.parent.moveCursor(QTextCursor.MoveOperation.Start)
        count = 0

        flags = QTextDocument.FindFlag(0)
        if self.match_case.isChecked():
            flags |= QTextDocument.FindCaseSensitively
        if self.whole_words.isChecked():
            flags |= QTextDocument.FindWholeWords

        while self.parent.find(self.find_edit.text(), flags):
            cursor = self.parent.textCursor()
            cursor.insertText(self.replace_edit.text())
            count += 1

        self.status_label.setText(f"Replaced {count} occurrences")


class DocumentMap(QWidget):
    def __init__(self, editor):
        super().__init__()
        self.editor = editor
        self.setup_ui()
        
    def setup_ui(self):
        self.setWindowTitle("Document Map")
        self.setMinimumWidth(200)
        
        self.layout = QVBoxLayout(self)
        self.list_widget = QListWidget()
        self.list_widget.itemClicked.connect(self.on_item_clicked)
        self.layout.addWidget(self.list_widget)
        
        # Update document map when content changes
        self.editor.document().contentsChange.connect(self.update_map)
        
    def update_map(self, position, chars_removed, chars_added):
        self.list_widget.clear()
        document = self.editor.document()
        block = document.begin()
        
        while block.isValid():
            if block.text().strip() and block.text().startswith('#'):
                self.list_widget.addItem(block.text().lstrip('#').strip())
            block = block.next()
    
    def on_item_clicked(self, item):
        # Find the clicked heading in the document
        document = self.editor.document()
        cursor = QTextCursor(document)
        
        while not cursor.atEnd():
            cursor.movePosition(QTextCursor.MoveOperation.NextBlock)
            if cursor.block().text().lstrip('#').strip() == item.text():
                self.editor.setTextCursor(cursor)
                self.editor.ensureCursorVisible()
                break


class GoToDialog(QDialog):
    def __init__(self, max_pages, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.max_pages = max_pages
        self.setup_ui()
        
    def setup_ui(self):
        self.setWindowTitle("Go To")
        layout = QVBoxLayout()
        
        self.page_edit = QLineEdit()
        self.page_edit.setPlaceholderText(f"Enter page number (1-{self.max_pages})")
        
        button_box = QHBoxLayout()
        go_btn = QPushButton("Go To")
        cancel_btn = QPushButton("Cancel")
        
        go_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        
        button_box.addWidget(go_btn)
        button_box.addWidget(cancel_btn)
        
        layout.addWidget(self.page_edit)
        layout.addLayout(button_box)
        
        self.setLayout(layout)
    
    def get_page_number(self):
        try:
            page = int(self.page_edit.text())
            if 1 <= page <= self.max_pages:
                return page
        except ValueError:
            pass
        return None
