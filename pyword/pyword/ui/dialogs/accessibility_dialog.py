"""
Accessibility Settings Dialog

This dialog allows users to configure accessibility features.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QGroupBox,
                              QCheckBox, QComboBox, QLabel, QPushButton,
                              QSpinBox, QDoubleSpinBox, QSlider, QTabWidget, QWidget)
from PySide6.QtCore import Qt


class AccessibilityDialog(QDialog):
    """Dialog for configuring accessibility settings."""

    def __init__(self, accessibility_manager, parent=None):
        super().__init__(parent)
        self.accessibility_manager = accessibility_manager
        self.setWindowTitle("Accessibility Settings")
        self.setMinimumSize(500, 400)
        self.setup_ui()
        self.load_current_settings()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout(self)

        # Create tab widget
        tab_widget = QTabWidget()

        # Screen Reader tab
        screen_reader_tab = self.create_screen_reader_tab()
        tab_widget.addTab(screen_reader_tab, "Screen Reader")

        # Keyboard Navigation tab
        keyboard_nav_tab = self.create_keyboard_nav_tab()
        tab_widget.addTab(keyboard_nav_tab, "Keyboard Navigation")

        # Visual tab
        visual_tab = self.create_visual_tab()
        tab_widget.addTab(visual_tab, "Visual")

        # Text-to-Speech tab
        tts_tab = self.create_tts_tab()
        tab_widget.addTab(tts_tab, "Text-to-Speech")

        layout.addWidget(tab_widget)

        # Buttons
        button_layout = QHBoxLayout()

        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)

        self.apply_button = QPushButton("Apply")
        self.apply_button.clicked.connect(self.apply_settings)

        button_layout.addStretch()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.apply_button)

        layout.addLayout(button_layout)

    def create_screen_reader_tab(self):
        """Create the screen reader settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Enable screen reader
        self.screen_reader_enabled = QCheckBox("Enable Screen Reader Support")
        layout.addWidget(self.screen_reader_enabled)

        # Announcement settings
        group = QGroupBox("Announcement Settings")
        group_layout = QVBoxLayout()

        self.announce_position = QCheckBox("Announce cursor position")
        self.announce_formatting = QCheckBox("Announce text formatting")
        self.announce_selection = QCheckBox("Announce text selection")

        group_layout.addWidget(self.announce_position)
        group_layout.addWidget(self.announce_formatting)
        group_layout.addWidget(self.announce_selection)

        group.setLayout(group_layout)
        layout.addWidget(group)

        layout.addStretch()
        return widget

    def create_keyboard_nav_tab(self):
        """Create the keyboard navigation settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Enable keyboard navigation
        self.keyboard_nav_enabled = QCheckBox("Enable Enhanced Keyboard Navigation")
        layout.addWidget(self.keyboard_nav_enabled)

        # Navigation options
        group = QGroupBox("Navigation Options")
        group_layout = QVBoxLayout()

        self.nav_by_paragraph = QCheckBox("Navigate by paragraph (Ctrl+Up/Down)")
        self.nav_by_heading = QCheckBox("Navigate by heading (Alt+Up/Down)")
        self.nav_by_element = QCheckBox("Navigate by element (Ctrl+Alt+T/I/L)")
        self.show_shortcut_hints = QCheckBox("Show keyboard shortcut hints")

        group_layout.addWidget(self.nav_by_paragraph)
        group_layout.addWidget(self.nav_by_heading)
        group_layout.addWidget(self.nav_by_element)
        group_layout.addWidget(self.show_shortcut_hints)

        group.setLayout(group_layout)
        layout.addWidget(group)

        layout.addStretch()
        return widget

    def create_visual_tab(self):
        """Create the visual settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # High contrast mode
        self.high_contrast_enabled = QCheckBox("Enable High Contrast Mode")
        layout.addWidget(self.high_contrast_enabled)

        # Theme selection
        theme_layout = QHBoxLayout()
        theme_layout.addWidget(QLabel("Theme:"))

        self.theme_combo = QComboBox()
        themes = self.accessibility_manager.high_contrast.get_available_themes()
        self.theme_combo.addItems(themes)
        theme_layout.addWidget(self.theme_combo)

        layout.addLayout(theme_layout)

        # Font size
        font_layout = QHBoxLayout()
        font_layout.addWidget(QLabel("Base Font Size:"))

        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(8, 72)
        self.font_size_spin.setValue(12)
        self.font_size_spin.setSuffix(" pt")
        font_layout.addWidget(self.font_size_spin)

        layout.addLayout(font_layout)

        # Cursor settings
        group = QGroupBox("Cursor Settings")
        group_layout = QVBoxLayout()

        self.large_cursor = QCheckBox("Use large cursor")
        self.blink_cursor = QCheckBox("Enable cursor blinking")

        group_layout.addWidget(self.large_cursor)
        group_layout.addWidget(self.blink_cursor)

        group.setLayout(group_layout)
        layout.addWidget(group)

        layout.addStretch()
        return widget

    def create_tts_tab(self):
        """Create the text-to-speech settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Enable TTS
        self.tts_enabled = QCheckBox("Enable Text-to-Speech")
        layout.addWidget(self.tts_enabled)

        # Voice selection
        voice_layout = QHBoxLayout()
        voice_layout.addWidget(QLabel("Voice:"))

        self.voice_combo = QComboBox()
        voices = self.accessibility_manager.tts.get_available_voices()
        self.voice_combo.addItems(voices)
        voice_layout.addWidget(self.voice_combo)

        layout.addLayout(voice_layout)

        # Speech rate
        rate_layout = QVBoxLayout()
        rate_layout.addWidget(QLabel("Speech Rate:"))

        rate_slider_layout = QHBoxLayout()
        self.rate_slider = QSlider(Qt.Horizontal)
        self.rate_slider.setRange(50, 200)
        self.rate_slider.setValue(100)
        self.rate_slider.setTickPosition(QSlider.TicksBelow)
        self.rate_slider.setTickInterval(25)

        self.rate_label = QLabel("1.0x")
        self.rate_slider.valueChanged.connect(
            lambda v: self.rate_label.setText(f"{v/100:.1f}x")
        )

        rate_slider_layout.addWidget(self.rate_slider)
        rate_slider_layout.addWidget(self.rate_label)

        rate_layout.addLayout(rate_slider_layout)
        layout.addLayout(rate_layout)

        # TTS Controls
        controls_group = QGroupBox("Controls")
        controls_layout = QVBoxLayout()

        self.read_selection_btn = QPushButton("Read Selected Text")
        self.read_selection_btn.clicked.connect(self.read_selection)

        self.stop_speech_btn = QPushButton("Stop Speech")
        self.stop_speech_btn.clicked.connect(self.stop_speech)

        controls_layout.addWidget(self.read_selection_btn)
        controls_layout.addWidget(self.stop_speech_btn)

        controls_group.setLayout(controls_layout)
        layout.addWidget(controls_group)

        layout.addStretch()
        return widget

    def load_current_settings(self):
        """Load current settings from accessibility manager."""
        # Screen reader
        self.screen_reader_enabled.setChecked(
            self.accessibility_manager.is_feature_enabled('screen_reader')
        )
        self.announce_position.setChecked(True)
        self.announce_formatting.setChecked(True)
        self.announce_selection.setChecked(True)

        # Keyboard navigation
        self.keyboard_nav_enabled.setChecked(
            self.accessibility_manager.is_feature_enabled('keyboard_nav')
        )
        self.nav_by_paragraph.setChecked(True)
        self.nav_by_heading.setChecked(True)
        self.nav_by_element.setChecked(True)
        self.show_shortcut_hints.setChecked(True)

        # Visual
        self.high_contrast_enabled.setChecked(
            self.accessibility_manager.is_feature_enabled('high_contrast')
        )
        current_theme = self.accessibility_manager.high_contrast.get_current_theme()
        index = self.theme_combo.findText(current_theme)
        if index >= 0:
            self.theme_combo.setCurrentIndex(index)

        # TTS
        self.tts_enabled.setChecked(
            self.accessibility_manager.is_feature_enabled('tts')
        )

    def apply_settings(self):
        """Apply the current settings."""
        # Screen reader
        if self.screen_reader_enabled.isChecked():
            self.accessibility_manager.enable_feature('screen_reader')
        else:
            self.accessibility_manager.disable_feature('screen_reader')

        # Keyboard navigation
        if self.keyboard_nav_enabled.isChecked():
            self.accessibility_manager.enable_feature('keyboard_nav')
        else:
            self.accessibility_manager.disable_feature('keyboard_nav')

        # High contrast
        if self.high_contrast_enabled.isChecked():
            self.accessibility_manager.enable_feature('high_contrast')
            theme = self.theme_combo.currentText()
            self.accessibility_manager.high_contrast.set_theme(theme)
        else:
            self.accessibility_manager.disable_feature('high_contrast')
            # Reset to default theme
            self.accessibility_manager.high_contrast.set_theme("default")

        # TTS
        if self.tts_enabled.isChecked():
            self.accessibility_manager.enable_feature('tts')
        else:
            self.accessibility_manager.disable_feature('tts')

        # Apply TTS settings
        rate = self.rate_slider.value() / 100.0
        self.accessibility_manager.tts.set_rate(rate)

        voice = self.voice_combo.currentText()
        if voice:
            self.accessibility_manager.tts.set_voice(voice)

    def accept(self):
        """Apply settings and close dialog."""
        self.apply_settings()
        super().accept()

    def read_selection(self):
        """Read selected text using TTS."""
        parent = self.parent()
        if parent and hasattr(parent, 'current_editor'):
            editor = parent.current_editor()
            if editor:
                self.accessibility_manager.tts.read_selection(editor)

    def stop_speech(self):
        """Stop text-to-speech."""
        self.accessibility_manager.tts.stop()
