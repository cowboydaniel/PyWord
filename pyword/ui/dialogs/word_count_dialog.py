from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                             QDialogButtonBox, QGroupBox, QGridLayout, QCheckBox)
from PySide6.QtCore import Qt

class WordCountDialog(QDialog):
    """Word count dialog showing document statistics."""
    
    def __init__(self, stats, parent=None):
        """Initialize the word count dialog.
        
        Args:
            stats: Dictionary containing document statistics
            parent: Parent widget
        """
        super().__init__(parent)
        self.setWindowTitle("Word Count")
        self.setFixedSize(400, 300)
        
        # Default statistics
        self.stats = {
            'pages': 1,
            'words': 0,
            'characters': 0,
            'characters_no_spaces': 0,
            'paragraphs': 0,
            'lines': 0,
            'footnotes': 0,
            'endnotes': 0,
            'text_boxes': 0,
            'in_footnotes': False
        }
        
        # Update with provided statistics
        if stats:
            self.stats.update(stats)
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the word count dialog UI."""
        layout = QVBoxLayout(self)
        
        # Statistics group
        stats_group = QGroupBox("Statistics")
        stats_layout = QGridLayout()
        
        # Add statistics rows
        self.add_stat_row(stats_layout, 0, "Pages:", self.stats['pages'])
        self.add_stat_row(stats_layout, 1, "Words:", self.stats['words'])
        self.add_stat_row(stats_layout, 2, "Characters (with spaces):", self.stats['characters'])
        self.add_stat_row(stats_layout, 3, "Characters (no spaces):", self.stats['characters_no_spaces'])
        self.add_stat_row(stats_layout, 4, "Paragraphs:", self.stats['paragraphs'])
        self.add_stat_row(stats_layout, 5, "Lines:", self.stats['lines'])
        
        # If in footnotes/endnotes, show additional stats
        if self.stats['in_footnotes']:
            self.add_stat_row(stats_layout, 6, "Footnotes:", self.stats['footnotes'])
            self.add_stat_row(stats_layout, 7, "Endnotes:", self.stats['endnotes'])
            self.add_stat_row(stats_layout, 8, "Text boxes:", self.stats['text_boxes'])
        
        stats_group.setLayout(stats_layout)
        
        # Include footnotes/endnotes checkbox (only show if not already in footnotes)
        if not self.stats['in_footnotes']:
            self.include_footnotes = QCheckBox("Include textboxes, footnotes, and endnotes")
            self.include_footnotes.setChecked(False)
            self.include_footnotes.toggled.connect(self.on_include_footnotes_toggled)
            
            layout.addWidget(self.include_footnotes)
            layout.addSpacing(10)
        
        layout.addWidget(stats_group, 1)
        
        # Close button
        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.rejected.connect(self.reject)
        
        layout.addWidget(button_box)
    
    def add_stat_row(self, layout, row, label, value):
        """Add a statistics row to the layout."""
        layout.addWidget(QLabel(label), row, 0, Qt.AlignLeft)
        layout.addWidget(QLabel(str(value)), row, 1, Qt.AlignRight)
    
    def on_include_footnotes_toggled(self, checked):
        """Handle include footnotes checkbox toggled."""
        # In a real implementation, this would update the statistics
        # to include/exclude footnotes and endnotes
        pass
    
    @staticmethod
    def show_dialog(stats=None, parent=None):
        """Static method to show the word count dialog."""
        dialog = WordCountDialog(stats, parent)
        dialog.exec()
