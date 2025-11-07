from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QCheckBox, QPushButton, QGroupBox, QDialog, QTabWidget,
                             QListWidget, QListWidgetItem, QMessageBox, QComboBox)
from PySide6.QtCore import Qt, Signal, QRegularExpression
from PySide6.QtGui import QTextCursor, QTextDocument, QTextCharFormat, QColor, QFont

class FindReplaceDialog(QDialog):
    """Find and replace dialog with advanced search options."""
    
    # Signals
    find_next = Signal(str, bool, bool, bool, bool)  # text, case_sensitive, whole_word, regex, wrap
    replace = Signal(str, str, bool, bool, bool, bool)  # find_text, replace_text, case_sensitive, whole_word, regex, replace_all
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Find and Replace")
        self.setMinimumSize(500, 400)
        
        self.find_history = []
        self.replace_history = []
        self.current_find_index = -1
        self.current_replace_index = -1
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the find/replace dialog UI."""
        layout = QVBoxLayout(self)
        
        # Create tabs
        self.tabs = QTabWidget()
        
        # Find tab
        find_tab = QWidget()
        self.setup_find_tab(find_tab)
        self.tabs.addTab(find_tab, "Find")
        
        # Replace tab
        replace_tab = QWidget()
        self.setup_replace_tab(replace_tab)
        self.tabs.addTab(replace_tab, "Replace")
        
        layout.addWidget(self.tabs)
        
        # Button box
        button_box = QDialogButtonBox()
        
        self.find_next_btn = button_box.addButton("Find Next", QDialogButtonBox.ActionRole)
        self.find_next_btn.setDefault(True)
        self.find_next_btn.clicked.connect(self.on_find_next)
        
        self.replace_btn = button_box.addButton("Replace", QDialogButtonBox.ActionRole)
        self.replace_btn.clicked.connect(self.on_replace)
        
        self.replace_all_btn = button_box.addButton("Replace All", QDialogButtonBox.ActionRole)
        self.replace_all_btn.clicked.connect(self.on_replace_all)
        
        close_btn = button_box.addButton("Close", QDialogButtonBox.RejectRole)
        close_btn.clicked.connect(self.reject)
        
        # Disable replace buttons by default (enabled when text is found)
        self.replace_btn.setEnabled(False)
        self.replace_all_btn.setEnabled(False)
        
        layout.addWidget(button_box)
    
    def setup_find_tab(self, parent):
        """Setup the Find tab."""
        layout = QVBoxLayout(parent)
        
        # Find what
        find_layout = QHBoxLayout()
        find_layout.addWidget(QLabel("Find what:"))
        
        self.find_combo = QComboBox()
        self.find_combo.setEditable(True)
        self.find_combo.setInsertPolicy(QComboBox.InsertAtTop)
        self.find_combo.setMaxCount(10)
        self.find_combo.setMaxVisibleItems(10)
        self.find_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.find_combo.editTextChanged.connect(self.update_find_buttons)
        self.find_combo.activated.connect(self.on_find_text_selected)
        
        find_layout.addWidget(self.find_combo)
        
        # Find next button
        self.find_next_tab_btn = QPushButton("Find Next")
        self.find_next_tab_btn.clicked.connect(self.on_find_next)
        
        find_layout.addWidget(self.find_next_tab_btn)
        
        layout.addLayout(find_layout)
        
        # Options group
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()
        
        # Match case
        self.match_case_check = QCheckBox("Match case")
        
        # Match whole word
        self.whole_word_check = QCheckBox("Match whole word only")
        
        # Use regular expressions
        self.regex_check = QCheckBox("Use regular expressions")
        
        # Wrap around
        self.wrap_check = QCheckBox("Wrap around")
        self.wrap_check.setChecked(True)
        
        options_layout.addWidget(self.match_case_check)
        options_layout.addWidget(self.whole_word_check)
        options_layout.addWidget(self.regex_check)
        options_layout.addWidget(self.wrap_check)
        options_group.setLayout(options_layout)
        
        layout.addWidget(options_group)
        
        # Direction
        direction_group = QGroupBox("Direction")
        direction_layout = QHBoxLayout()
        
        self.direction_up = QRadioButton("Up")
        self.direction_down = QRadioButton("Down")
        self.direction_down.setChecked(True)
        
        direction_layout.addWidget(self.direction_up)
        direction_layout.addWidget(self.direction_down)
        direction_group.setLayout(direction_layout)
        
        layout.addWidget(direction_group)
        layout.addStretch()
    
    def setup_replace_tab(self, parent):
        """Setup the Replace tab."""
        layout = QVBoxLayout(parent)
        
        # Find what
        find_layout = QHBoxLayout()
        find_layout.addWidget(QLabel("Find what:"))
        
        self.replace_find_combo = QComboBox()
        self.replace_find_combo.setEditable(True)
        self.replace_find_combo.setInsertPolicy(QComboBox.InsertAtTop)
        self.replace_find_combo.setMaxCount(10)
        self.replace_find_combo.setMaxVisibleItems(10)
        self.replace_find_combo.editTextChanged.connect(self.update_find_buttons)
        self.replace_find_combo.activated.connect(self.on_find_text_selected)
        
        find_layout.addWidget(self.replace_find_combo)
        layout.addLayout(find_layout)
        
        # Replace with
        replace_layout = QHBoxLayout()
        replace_layout.addWidget(QLabel("Replace with:"))
        
        self.replace_combo = QComboBox()
        self.replace_combo.setEditable(True)
        self.replace_combo.setInsertPolicy(QComboBox.InsertAtTop)
        self.replace_combo.setMaxCount(10)
        self.replace_combo.setMaxVisibleItems(10)
        self.replace_combo.activated.connect(self.on_replace_text_selected)
        
        replace_layout.addWidget(self.replace_combo)
        layout.addLayout(replace_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.replace_btn_tab = QPushButton("Replace")
        self.replace_btn_tab.clicked.connect(self.on_replace)
        
        self.replace_all_btn_tab = QPushButton("Replace All")
        self.replace_all_btn_tab.clicked.connect(self.on_replace_all)
        
        button_layout.addWidget(self.replace_btn_tab)
        button_layout.addWidget(self.replace_all_btn_tab)
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        
        # Options group (same as Find tab)
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()
        
        # Match case
        self.replace_match_case_check = QCheckBox("Match case")
        self.replace_match_case_check.setChecked(self.match_case_check.isChecked())
        self.match_case_check.toggled.connect(self.replace_match_case_check.setChecked)
        self.replace_match_case_check.toggled.connect(self.match_case_check.setChecked)
        
        # Match whole word
        self.replace_whole_word_check = QCheckBox("Match whole word only")
        self.replace_whole_word_check.setChecked(self.whole_word_check.isChecked())
        self.whole_word_check.toggled.connect(self.replace_whole_word_check.setChecked)
        self.replace_whole_word_check.toggled.connect(self.whole_word_check.setChecked)
        
        # Use regular expressions
        self.replace_regex_check = QCheckBox("Use regular expressions")
        self.replace_regex_check.setChecked(self.regex_check.isChecked())
        self.regex_check.toggled.connect(self.replace_regex_check.setChecked)
        self.replace_regex_check.toggled.connect(self.regex_check.setChecked)
        
        # Wrap around
        self.replace_wrap_check = QCheckBox("Wrap around")
        self.replace_wrap_check.setChecked(self.wrap_check.isChecked())
        self.wrap_check.toggled.connect(self.replace_wrap_check.setChecked)
        self.replace_wrap_check.toggled.connect(self.wrap_check.setChecked)
        
        options_layout.addWidget(self.replace_match_case_check)
        options_layout.addWidget(self.replace_whole_word_check)
        options_layout.addWidget(self.replace_regex_check)
        options_layout.addWidget(self.replace_wrap_check)
        options_group.setLayout(options_layout)
        
        layout.addWidget(options_group)
        
        # Direction
        direction_group = QGroupBox("Direction")
        direction_layout = QHBoxLayout()
        
        self.replace_direction_up = QRadioButton("Up")
        self.replace_direction_down = QRadioButton("Down")
        self.replace_direction_down.setChecked(True)
        
        # Sync direction between tabs
        self.direction_up.toggled.connect(self.replace_direction_up.setChecked)
        self.direction_down.toggled.connect(self.replace_direction_down.setChecked)
        self.replace_direction_up.toggled.connect(self.direction_up.setChecked)
        self.replace_direction_down.toggled.connect(self.direction_down.setChecked)
        
        direction_layout.addWidget(self.replace_direction_up)
        direction_layout.addWidget(self.replace_direction_down)
        direction_group.setLayout(direction_layout)
        
        layout.addWidget(direction_group)
        layout.addStretch()
    
    def update_find_buttons(self, text):
        """Update the state of find/replace buttons based on input."""
        has_text = bool(text.strip())
        
        self.find_next_btn.setEnabled(has_text)
        self.find_next_tab_btn.setEnabled(has_text)
        self.replace_btn.setEnabled(has_text)
        self.replace_btn_tab.setEnabled(has_text)
        self.replace_all_btn.setEnabled(has_text)
        self.replace_all_btn_tab.setEnabled(has_text)
    
    def on_find_text_selected(self, index):
        """Handle selection from find history."""
        if index >= 0 and index < len(self.find_history):
            self.current_find_index = index
            self.find_combo.setEditText(self.find_history[index])
            self.replace_find_combo.setEditText(self.find_history[index])
    
    def on_replace_text_selected(self, index):
        """Handle selection from replace history."""
        if index >= 0 and index < len(self.replace_history):
            self.current_replace_index = index
            self.replace_combo.setEditText(self.replace_history[index])
    
    def on_find_next(self):
        """Handle Find Next button click."""
        find_text = self.find_combo.currentText()
        
        if not find_text:
            return
        
        # Add to history if not already present
        if find_text not in self.find_history:
            self.find_history.insert(0, find_text)
            self.find_combo.insertItem(0, find_text)
            self.replace_find_combo.insertItem(0, find_text)
            
            # Limit history size
            if len(self.find_history) > 10:
                self.find_history.pop()
                self.find_combo.removeItem(self.find_combo.count() - 1)
                self.replace_find_combo.removeItem(self.replace_find_combo.count() - 1)
        
        # Emit signal with current options
        self.find_next.emit(
            find_text,
            self.match_case_check.isChecked(),
            self.whole_word_check.isChecked(),
            self.regex_check.isChecked(),
            self.wrap_check.isChecked()
        )
    
    def on_replace(self):
        """Handle Replace button click."""
        find_text = self.replace_find_combo.currentText()
        replace_text = self.replace_combo.currentText()
        
        if not find_text:
            return
        
        # Add to history if not already present
        if find_text not in self.find_history:
            self.find_history.insert(0, find_text)
            self.find_combo.insertItem(0, find_text)
            self.replace_find_combo.insertItem(0, find_text)
            
            if len(self.find_history) > 10:
                self.find_history.pop()
                self.find_combo.removeItem(self.find_combo.count() - 1)
                self.replace_find_combo.removeItem(self.replace_find_combo.count() - 1)
        
        if replace_text not in self.replace_history:
            self.replace_history.insert(0, replace_text)
            self.replace_combo.insertItem(0, replace_text)
            
            if len(self.replace_history) > 10:
                self.replace_history.pop()
                self.replace_combo.removeItem(self.replace_combo.count() - 1)
        
        # Emit signal with current options
        self.replace.emit(
            find_text,
            replace_text,
            self.replace_match_case_check.isChecked(),
            self.replace_whole_word_check.isChecked(),
            self.replace_regex_check.isChecked(),
            False  # replace_all
        )
    
    def on_replace_all(self):
        """Handle Replace All button click."""
        find_text = self.replace_find_combo.currentText()
        replace_text = self.replace_combo.currentText()
        
        if not find_text:
            return
        
        # Add to history if not already present
        if find_text not in self.find_history:
            self.find_history.insert(0, find_text)
            self.find_combo.insertItem(0, find_text)
            self.replace_find_combo.insertItem(0, find_text)
            
            if len(self.find_history) > 10:
                self.find_history.pop()
                self.find_combo.removeItem(self.find_combo.count() - 1)
                self.replace_find_combo.removeItem(self.replace_find_combo.count() - 1)
        
        if replace_text not in self.replace_history:
            self.replace_history.insert(0, replace_text)
            self.replace_combo.insertItem(0, replace_text)
            
            if len(self.replace_history) > 10:
                self.replace_history.pop()
                self.replace_combo.removeItem(self.replace_combo.count() - 1)
        
        # Ask for confirmation if replacing many occurrences
        if find_text:
            reply = QMessageBox.question(
                self,
                "Replace All",
                f"Replace all occurrences of \"{find_text}\" with \"{replace_text}\"?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                # Emit signal with current options
                self.replace.emit(
                    find_text,
                    replace_text,
                    self.replace_match_case_check.isChecked(),
                    self.replace_whole_word_check.isChecked(),
                    self.replace_regex_check.isChecked(),
                    True  # replace_all
                )
    
    def show_find(self):
        """Show the find tab and set focus to the find field."""
        self.tabs.setCurrentIndex(0)
        self.find_combo.setFocus()
        self.find_combo.selectAll()
        self.show()
        self.activateWindow()
    
    def show_replace(self):
        """Show the replace tab and set focus to the find field."""
        self.tabs.setCurrentIndex(1)
        self.replace_find_combo.setFocus()
        self.replace_find_combo.selectAll()
        self.show()
        self.activateWindow()
    
    def set_find_text(self, text):
        """Set the find text."""
        self.find_combo.setEditText(text)
        self.replace_find_combo.setEditText(text)
    
    def set_replace_text(self, text):
        """Set the replace text."""
        self.replace_combo.setEditText(text)
    
    def get_find_text(self):
        """Get the current find text."""
        return self.find_combo.currentText()
    
    def get_replace_text(self):
        """Get the current replace text."""
        return self.replace_combo.currentText()
    
    def get_options(self):
        """Get the current search options."""
        return {
            'match_case': self.match_case_check.isChecked(),
            'whole_word': self.whole_word_check.isChecked(),
            'regex': self.regex_check.isChecked(),
            'wrap': self.wrap_check.isChecked(),
            'direction': 'up' if self.direction_up.isChecked() else 'down'
        }
    
    def set_options(self, options):
        """Set the search options."""
        if 'match_case' in options:
            self.match_case_check.setChecked(options['match_case'])
            self.replace_match_case_check.setChecked(options['match_case'])
        
        if 'whole_word' in options:
            self.whole_word_check.setChecked(options['whole_word'])
            self.replace_whole_word_check.setChecked(options['whole_word'])
        
        if 'regex' in options:
            self.regex_check.setChecked(options['regex'])
            self.replace_regex_check.setChecked(options['regex'])
        
        if 'wrap' in options:
            self.wrap_check.setChecked(options['wrap'])
            self.replace_wrap_check.setChecked(options['wrap'])
        
        if 'direction' in options:
            if options['direction'].lower() == 'up':
                self.direction_up.setChecked(True)
                self.replace_direction_up.setChecked(True)
            else:
                self.direction_down.setChecked(True)
                self.replace_direction_down.setChecked(True)
