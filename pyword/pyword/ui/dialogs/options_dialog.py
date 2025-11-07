from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                             QDialogButtonBox, QTabWidget, QWidget, QListWidget, QStackedWidget,
                             QCheckBox, QComboBox, QSpinBox, QDoubleSpinBox, QLineEdit, QGroupBox,
                             QFormLayout, QFileDialog, QColorDialog, QFontComboBox, QScrollArea)
from PySide6.QtCore import Qt, QSettings, Signal
from PySide6.QtGui import QColor, QFont

class OptionsDialog(QDialog):
    """Options dialog for application settings."""
    
    # Signal emitted when settings are changed
    settings_changed = Signal()
    
    def __init__(self, parent=None):
        """Initialize the options dialog."""
        super().__init__(parent)
        self.setWindowTitle("Options")
        self.setMinimumSize(700, 500)
        
        # Load settings
        self.settings = QSettings("PyWord", "PyWord")
        
        # Default settings
        self.default_settings = {
            'general/auto_recovery': True,
            'general/auto_save': True,
            'general/auto_save_interval': 10,  # minutes
            'general/recent_files': 10,
            'general/show_status_bar': True,
            'general/show_rulers': True,
            'general/show_hidden_chars': False,
            'general/measurement_units': 'cm',  # cm, mm, in, pt, pica
            
            'view/zoom': 100,
            'view/show_paragraph_marks': True,
            'view/show_hidden_text': False,
            'view/show_bookmarks': True,
            'view/show_text_boundaries': True,
            'view/show_highlight': True,
            
            'edit/typing_replaces_selection': True,
            'edit/use_smart_cut_paste': True,
            'edit/allow_accents': True,
            'edit/enable_click_and_type': True,
            'edit/default_paragraph_style': 'Normal',
            'edit/default_font_family': 'Arial',
            'edit/default_font_size': 12,
            'edit/auto_format_as_you_type': True,
            'edit/auto_correct': True,
            'edit/auto_text': True,
            
            'save/auto_recover': True,
            'save/always_create_backup': False,
            'save/embed_fonts': False,
            'save/embed_system_fonts': False,
            'save/do_not_embed_common_system_fonts': True,
            
            'spelling/check_spelling_as_you_type': True,
            'spelling/highlight_errors': True,
            'spelling/ignore_uppercase': True,
            'spelling/ignore_words_with_numbers': True,
            'spelling/ignore_internet_and_file_addresses': True,
            'spelling/suggest_from_main_dictionary_only': False,
            'spelling/custom_dictionary': '',
            
            'advanced/confirm_conversion_at_open': True,
            'advanced/update_automatic_links_at_open': True,
            'advanced/mail_as_author': False,
            'advanced/store_random_number_to_combine_merges': False,
            'advanced/use_character_units': False,
            'advanced/pixels_per_inch': 96,
            'advanced/show_all_windows_in_taskbar': True,
            'advanced/use_subpixel_positioning': True,
            'advanced/optimize_character_positioning_for_layout': False,
            'advanced/do_not_use_printer_metrics': False,
            'advanced/use_legacy_track_changes': False,
            'advanced/allow_open_in_draft_view': False,
            'advanced/allow_background_open_web_pages': False,
            'advanced/allow_clipboard_notification': True,
            'advanced/allow_drag_and_drop': True,
            'advanced/show_shortcut_keys_in_screen_tips': True,
            'advanced/show_vertical_ruler': True,
            'advanced/use_smart_paragraph_selection': True,
            'advanced/use_smart_cursor_movement': True,
            'advanced/use_ins_key_for_paste': False,
            'advanced/overtype_mode': False,
            'advanced/use_normal_style_for_list': False,
            'advanced/use_character_units': False,
            'advanced/use_legacy_track_changes': False,
            'advanced/allow_click_and_type': True,
            'advanced/use_subpixel_positioning': True
        }
        
        # Create settings if they don't exist
        for key, value in self.default_settings.items():
            if not self.settings.contains(key):
                self.settings.setValue(key, value)
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the options dialog UI."""
        main_layout = QHBoxLayout(self)
        
        # Left side: Categories list
        self.categories = QListWidget()
        self.categories.setFixedWidth(150)
        self.categories.setIconSize(QSize(16, 16))
        self.categories.currentRowChanged.connect(self.change_page)
        
        # Right side: Stacked widget for settings pages
        self.pages = QStackedWidget()
        
        # Create pages
        self.create_general_page()
        self.create_display_page()
        self.create_edit_page()
        self.create_save_page()
        self.create_spelling_page()
        self.create_advanced_page()
        
        # Add categories to list
        categories = ["General", "Display", "Edit", "Save", "Spelling", "Advanced"]
        self.categories.addItems(categories)
        
        # Set first category as current
        self.categories.setCurrentRow(0)
        
        # Add widgets to main layout
        main_layout.addWidget(self.categories)
        main_layout.addWidget(self.pages, 1)
        
        # Button box
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Apply | QDialogButtonBox.Help | QDialogButtonBox.RestoreDefaults
        )
        
        # Connect buttons
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        apply_btn = button_box.button(QDialogButtonBox.Apply)
        apply_btn.clicked.connect(self.apply_changes)
        
        defaults_btn = button_box.button(QDialogButtonBox.RestoreDefaults)
        defaults_btn.clicked.connect(self.restore_defaults)
        
        # Add button box to layout
        main_layout.addWidget(button_box, 0, Qt.AlignBottom)
    
    def create_general_page(self):
        """Create the General options page."""
        page = QWidget()
        layout = QVBoxLayout(page)
        
        # Auto-recover group
        recover_group = QGroupBox("AutoRecover")
        recover_layout = QVBoxLayout()
        
        self.auto_recover = QCheckBox("Save AutoRecover information every")
        self.auto_recover.setChecked(self.settings.value('general/auto_recovery', True, type=bool))
        
        self.auto_save_interval = QSpinBox()
        self.auto_save_interval.setRange(1, 120)
        self.auto_save_interval.setValue(self.settings.value('general/auto_save_interval', 10, type=int))
        self.auto_save_interval.setSuffix(" minutes")
        
        recover_layout.addWidget(self.auto_recover)
        
        interval_layout = QHBoxLayout()
        interval_layout.addSpacing(20)
        interval_layout.addWidget(QLabel("Save every:"))
        interval_layout.addWidget(self.auto_save_interval)
        interval_layout.addStretch()
        
        recover_layout.addLayout(interval_layout)
        recover_group.setLayout(recover_layout)
        
        # Recent files group
        recent_group = QGroupBox("Recent Files")
        recent_layout = QVBoxLayout()
        
        self.show_recent_files = QCheckBox("Quickly access this number of Recent Documents:")
        self.show_recent_files.setChecked(True)
        
        self.recent_files_count = QSpinBox()
        self.recent_files_count.setRange(0, 50)
        self.recent_files_count.setValue(self.settings.value('general/recent_files', 10, type=int))
        
        recent_layout.addWidget(self.show_recent_files)
        
        count_layout = QHBoxLayout()
        count_layout.addSpacing(20)
        count_layout.addWidget(QLabel("Number of documents:"))
        count_layout.addWidget(self.recent_files_count)
        count_layout.addStretch()
        
        recent_layout.addLayout(count_layout)
        recent_group.setLayout(recent_layout)
        
        # Add groups to layout
        layout.addWidget(recover_group)
        layout.addWidget(recent_group)
        layout.addStretch()
        
        # Add page to stacked widget
        self.pages.addWidget(page)
    
    def create_display_page(self):
        """Create the Display options page."""
        page = QWidget()
        layout = QVBoxLayout(page)
        
        # Page display options group
        display_group = QGroupBox("Page Display Options")
        display_layout = QVBoxLayout()
        
        self.show_status_bar = QCheckBox("Show status bar")
        self.show_status_bar.setChecked(self.settings.value('general/show_status_bar', True, type=bool))
        
        self.show_rulers = QCheckBox("Show vertical ruler")
        self.show_rulers.setChecked(self.settings.value('general/show_rulers', True, type=bool))
        
        self.show_paragraph_marks = QCheckBox("Show paragraph marks")
        self.show_paragraph_marks.setChecked(self.settings.value('view/show_paragraph_marks', True, type=bool))
        
        self.show_hidden_text = QCheckBox("Show hidden text")
        self.show_hidden_text.setChecked(self.settings.value('view/show_hidden_text', False, type=bool))
        
        self.show_bookmarks = QCheckBox("Show bookmarks")
        self.show_bookmarks.setChecked(self.settings.value('view/show_bookmarks', True, type=bool))
        
        display_layout.addWidget(self.show_status_bar)
        display_layout.addWidget(self.show_rulers)
        display_layout.addWidget(self.show_paragraph_marks)
        display_layout.addWidget(self.show_hidden_text)
        display_layout.addWidget(self.show_bookmarks)
        
        display_group.setLayout(display_layout)
        
        # Zoom group
        zoom_group = QGroupBox("Zoom")
        zoom_layout = QVBoxLayout()
        
        self.default_zoom = QSpinBox()
        self.default_zoom.setRange(10, 500)
        self.default_zoom.setValue(self.settings.value('view/zoom', 100, type=int))
        self.default_zoom.setSuffix("%")
        
        zoom_layout.addWidget(QLabel("Default zoom percentage:"))
        zoom_layout.addWidget(self.default_zoom)
        zoom_group.setLayout(zoom_layout)
        
        # Add groups to layout
        layout.addWidget(display_group)
        layout.addWidget(zoom_group)
        layout.addStretch()
        
        # Add page to stacked widget
        self.pages.addWidget(page)
    
    def create_edit_page(self):
        """Create the Edit options page."""
        page = QWidget()
        layout = QVBoxLayout(page)
        
        # Editing options group
        edit_group = QGroupBox("Editing Options")
        edit_layout = QVBoxLayout()
        
        self.typing_replaces_selection = QCheckBox("Typing replaces selection")
        self.typing_replaces_selection.setChecked(self.settings.value('edit/typing_replaces_selection', True, type=bool))
        
        self.use_smart_cut_paste = QCheckBox("Use smart cut and paste")
        self.use_smart_cut_paste.setChecked(self.settings.value('edit/use_smart_cut_paste', True, type=bool))
        
        self.auto_format_as_you_type = QCheckBox("AutoFormat As You Type")
        self.auto_format_as_you_type.setChecked(self.settings.value('edit/auto_format_as_you_type', True, type=bool))
        
        self.auto_correct = QCheckBox("AutoCorrect")
        self.auto_correct.setChecked(self.settings.value('edit/auto_correct', True, type=bool))
        
        edit_layout.addWidget(self.typing_replaces_selection)
        edit_layout.addWidget(self.use_smart_cut_paste)
        edit_layout.addWidget(self.auto_format_as_you_type)
        edit_layout.addWidget(self.auto_correct)
        
        edit_group.setLayout(edit_layout)
        
        # Default font group
        font_group = QGroupBox("Default Font")
        font_layout = QFormLayout()
        
        self.default_font_family = QFontComboBox()
        self.default_font_family.setCurrentFont(QFont(self.settings.value('edit/default_font_family', 'Arial')))
        
        self.default_font_size = QSpinBox()
        self.default_font_size.setRange(1, 1638)
        self.default_font_size.setValue(self.settings.value('edit/default_font_size', 12, type=int))
        
        font_layout.addRow("Font:", self.default_font_family)
        font_layout.addRow("Size:", self.default_font_size)
        
        font_group.setLayout(font_layout)
        
        # Add groups to layout
        layout.addWidget(edit_group)
        layout.addWidget(font_group)
        layout.addStretch()
        
        # Add page to stacked widget
        self.pages.addWidget(page)
    
    def create_save_page(self):
        """Create the Save options page."""
        page = QWidget()
        layout = QVBoxLayout(page)
        
        # Save options group
        save_group = QGroupBox("Save Options")
        save_layout = QVBoxLayout()
        
        self.always_create_backup = QCheckBox("Always create backup copy")
        self.always_create_backup.setChecked(self.settings.value('save/always_create_backup', False, type=bool))
        
        self.embed_fonts = QCheckBox("Embed fonts in the file")
        self.embed_fonts.setChecked(self.settings.value('save/embed_fonts', False, type=bool))
        
        save_layout.addWidget(self.always_create_backup)
        save_layout.addWidget(self.embed_fonts)
        
        save_group.setLayout(save_layout)
        
        # File locations group
        locations_group = QGroupBox("File Locations")
        locations_layout = QFormLayout()
        
        self.auto_recover_location = QLineEdit(self.settings.value('save/auto_recover_location', ''))
        self.auto_recover_browse = QPushButton("Browse...")
        self.auto_recover_browse.clicked.connect(lambda: self.browse_directory(self.auto_recover_location))
        
        locations_layout.addRow("AutoRecover files location:", self.auto_recover_location)
        locations_layout.addRow("", self.auto_recover_browse)
        
        locations_group.setLayout(locations_layout)
        
        # Add groups to layout
        layout.addWidget(save_group)
        layout.addWidget(locations_group)
        layout.addStretch()
        
        # Add page to stacked widget
        self.pages.addWidget(page)
    
    def create_spelling_page(self):
        """Create the Spelling & Grammar options page."""
        page = QWidget()
        layout = QVBoxLayout(page)
        
        # Spelling options group
        spelling_group = QGroupBox("Spelling")
        spelling_layout = QVBoxLayout()
        
        self.check_spelling = QCheckBox("Check spelling as you type")
        self.check_spelling.setChecked(self.settings.value('spelling/check_spelling_as_you_type', True, type=bool))
        
        self.highlight_errors = QCheckBox("Frequently confused words")
        self.highlight_errors.setChecked(self.settings.value('spelling/highlight_errors', True, type=bool))
        
        self.ignore_uppercase = QCheckBox("Ignore words in UPPERCASE")
        self.ignore_uppercase.setChecked(self.settings.value('spelling/ignore_uppercase', True, type=bool))
        
        self.ignore_numbers = QCheckBox("Ignore words with numbers")
        self.ignore_numbers.setChecked(self.settings.value('spelling/ignore_words_with_numbers', True, type=bool))
        
        self.ignore_urls = QCheckBox("Ignore Internet and file addresses")
        self.ignore_urls.setChecked(self.settings.value('spelling/ignore_internet_and_file_addresses', True, type=bool))
        
        spelling_layout.addWidget(self.check_spelling)
        spelling_layout.addWidget(self.highlight_errors)
        spelling_layout.addWidget(self.ignore_uppercase)
        spelling_layout.addWidget(self.ignore_numbers)
        spelling_layout.addWidget(self.ignore_urls)
        
        spelling_group.setLayout(spelling_layout)
        
        # Custom dictionary group
        dict_group = QGroupBox("Custom Dictionaries")
        dict_layout = QVBoxLayout()
        
        self.custom_dict = QLineEdit(self.settings.value('spelling/custom_dictionary', ''))
        self.custom_dict_browse = QPushButton("Browse...")
        self.custom_dict_browse.clicked.connect(lambda: self.browse_file(self.custom_dict, "Dictionary Files (*.dic)"))
        
        dict_layout.addWidget(QLabel("Custom dictionary:"))
        dict_layout.addWidget(self.custom_dict)
        dict_layout.addWidget(self.custom_dict_browse)
        
        dict_group.setLayout(dict_layout)
        
        # Add groups to layout
        layout.addWidget(spelling_group)
        layout.addWidget(dict_group)
        layout.addStretch()
        
        # Add page to stacked widget
        self.pages.addWidget(page)
    
    def create_advanced_page(self):
        """Create the Advanced options page."""
        page = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(page)
        
        layout = QVBoxLayout(page)
        
        # Editing options group
        edit_group = QGroupBox("Editing Options")
        edit_layout = QVBoxLayout()
        
        self.use_smart_cursor = QCheckBox("Use smart cursor movement")
        self.use_smart_cursor.setChecked(self.settings.value('advanced/use_smart_cursor_movement', True, type=bool))
        
        self.use_smart_paragraph = QCheckBox("Use smart paragraph selection")
        self.use_smart_paragraph.setChecked(self.settings.value('advanced/use_smart_paragraph_selection', True, type=bool))
        
        self.allow_drag_drop = QCheckBox("Allow text to be dragged and dropped")
        self.allow_drag_drop.setChecked(self.settings.value('advanced/allow_drag_and_drop', True, type=bool))
        
        self.use_ins_key = QCheckBox("Use the Insert key to control overtype mode")
        self.use_ins_key.setChecked(self.settings.value('advanced/use_ins_key_for_paste', False, type=bool))
        
        self.overtype_mode = QCheckBox("Use overtype mode")
        self.overtype_mode.setChecked(self.settings.value('advanced/overtype_mode', False, type=bool))
        
        edit_layout.addWidget(self.use_smart_cursor)
        edit_layout.addWidget(self.use_smart_paragraph)
        edit_layout.addWidget(self.allow_drag_drop)
        edit_layout.addWidget(self.use_ins_key)
        edit_layout.addWidget(self.overtype_mode)
        
        edit_group.setLayout(edit_layout)
        
        # Display group
        display_group = QGroupBox("Display")
        display_layout = QVBoxLayout()
        
        self.show_shortcut_keys = QCheckBox("Show shortcut keys in ScreenTips")
        self.show_shortcut_keys.setChecked(self.settings.value('advanced/show_shortcut_keys_in_screen_tips', True, type=bool))
        
        self.show_vertical_ruler = QCheckBox("Show vertical ruler")
        self.show_vertical_ruler.setChecked(self.settings.value('advanced/show_vertical_ruler', True, type=bool))
        
        display_layout.addWidget(self.show_shortcut_keys)
        display_layout.addWidget(self.show_vertical_ruler)
        
        display_group.setLayout(display_layout)
        
        # Add groups to layout
        layout.addWidget(edit_group)
        layout.addWidget(display_group)
        layout.addStretch()
        
        # Add page to stacked widget
        self.pages.addWidget(scroll)
    
    def browse_directory(self, line_edit):
        """Open a directory selection dialog."""
        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            line_edit.setText(directory)
    
    def browse_file(self, line_edit, file_filter):
        """Open a file selection dialog."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", file_filter)
        if file_path:
            line_edit.setText(file_path)
    
    def change_page(self, index):
        """Change the current page in the stacked widget."""
        self.pages.setCurrentIndex(index)
    
    def accept(self):
        """Handle dialog acceptance."""
        self.apply_changes()
        super().accept()
    
    def apply_changes(self):
        """Apply changes to settings."""
        # General tab
        self.settings.setValue('general/auto_recovery', self.auto_recover.isChecked())
        self.settings.setValue('general/auto_save_interval', self.auto_save_interval.value())
        self.settings.setValue('general/recent_files', self.recent_files_count.value())
        
        # Display tab
        self.settings.setValue('general/show_status_bar', self.show_status_bar.isChecked())
        self.settings.setValue('general/show_rulers', self.show_rulers.isChecked())
        self.settings.setValue('view/show_paragraph_marks', self.show_paragraph_marks.isChecked())
        self.settings.setValue('view/show_hidden_text', self.show_hidden_text.isChecked())
        self.settings.setValue('view/show_bookmarks', self.show_bookmarks.isChecked())
        self.settings.setValue('view/zoom', self.default_zoom.value())
        
        # Edit tab
        self.settings.setValue('edit/typing_replaces_selection', self.typing_replaces_selection.isChecked())
        self.settings.setValue('edit/use_smart_cut_paste', self.use_smart_cut_paste.isChecked())
        self.settings.setValue('edit/auto_format_as_you_type', self.auto_format_as_you_type.isChecked())
        self.settings.setValue('edit/auto_correct', self.auto_correct.isChecked())
        self.settings.setValue('edit/default_font_family', self.default_font_family.currentFont().family())
        self.settings.setValue('edit/default_font_size', self.default_font_size.value())
        
        # Save tab
        self.settings.setValue('save/always_create_backup', self.always_create_backup.isChecked())
        self.settings.setValue('save/embed_fonts', self.embed_fonts.isChecked())
        self.settings.setValue('save/auto_recover_location', self.auto_recover_location.text())
        
        # Spelling tab
        self.settings.setValue('spelling/check_spelling_as_you_type', self.check_spelling.isChecked())
        self.settings.setValue('spelling/highlight_errors', self.highlight_errors.isChecked())
        self.settings.setValue('spelling/ignore_uppercase', self.ignore_uppercase.isChecked())
        self.settings.setValue('spelling/ignore_words_with_numbers', self.ignore_numbers.isChecked())
        self.settings.setValue('spelling/ignore_internet_and_file_addresses', self.ignore_urls.isChecked())
        self.settings.setValue('spelling/custom_dictionary', self.custom_dict.text())
        
        # Advanced tab
        self.settings.setValue('advanced/use_smart_cursor_movement', self.use_smart_cursor.isChecked())
        self.settings.setValue('advanced/use_smart_paragraph_selection', self.use_smart_paragraph.isChecked())
        self.settings.setValue('advanced/allow_drag_and_drop', self.allow_drag_drop.isChecked())
        self.settings.setValue('advanced/use_ins_key_for_paste', self.use_ins_key.isChecked())
        self.settings.setValue('advanced/overtype_mode', self.overtype_mode.isChecked())
        self.settings.setValue('advanced/show_shortcut_keys_in_screen_tips', self.show_shortcut_keys.isChecked())
        self.settings.setValue('advanced/show_vertical_ruler', self.show_vertical_ruler.isChecked())
        
        # Emit signal that settings have changed
        self.settings_changed.emit()
    
    def restore_defaults(self):
        """Restore all settings to their default values."""
        # Reset all settings to defaults
        for key, value in self.default_settings.items():
            self.settings.setValue(key, value)
        
        # Update UI to reflect default values
        self.setup_ui()
        
        # Emit signal that settings have changed
        self.settings_changed.emit()
    
    @staticmethod
    def show_dialog(parent=None):
        """Static method to show the options dialog."""
        dialog = OptionsDialog(parent)
        return dialog.exec() == QDialog.Accepted
