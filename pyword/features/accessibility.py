"""
Accessibility Features for PyWord

This module provides accessibility features including:
- Screen reader support
- Keyboard navigation
- High contrast mode
- Text-to-speech
"""

import os
from typing import Optional, Dict, Any, List, Callable
from enum import Enum, auto
from PySide6.QtCore import QObject, Qt, Signal, QEvent, QTimer
from PySide6.QtGui import QPalette, QColor, QKeySequence, QTextCursor, QAccessible, QShortcut
from PySide6.QtWidgets import QWidget, QTextEdit, QApplication


class AccessibilityLevel(Enum):
    """Accessibility support levels."""
    NONE = auto()
    BASIC = auto()
    FULL = auto()


class ScreenReaderSupport(QObject):
    """
    Provides screen reader support for the application.

    Features:
    - ARIA-like attribute management
    - Focus announcements
    - Document structure navigation
    - Status updates
    """

    announcement_requested = Signal(str, int)  # message, priority

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.enabled = True
        self.current_focus = None
        self.document_landmarks: List[Dict[str, Any]] = []

    def enable(self):
        """Enable screen reader support."""
        self.enabled = True
        self.announce("Screen reader support enabled", priority=1)

    def disable(self):
        """Disable screen reader support."""
        self.enabled = False

    def announce(self, message: str, priority: int = 0):
        """
        Announce a message to the screen reader.

        Args:
            message: Message to announce
            priority: Priority level (0=low, 1=medium, 2=high)
        """
        if self.enabled:
            self.announcement_requested.emit(message, priority)
            # Update accessible description
            if self.parent():
                QAccessible.updateAccessibility(
                    self.parent(),
                    0,
                    QAccessible.Event.DescriptionChanged
                )

    def set_focus_announcement(self, widget: QWidget, description: str):
        """
        Set announcement for when a widget receives focus.

        Args:
            widget: Widget to monitor
            description: Description to announce
        """
        if widget:
            widget.setAccessibleName(description)
            widget.setAccessibleDescription(description)

    def announce_document_position(self, line: int, column: int, total_lines: int):
        """
        Announce current document position.

        Args:
            line: Current line number
            column: Current column number
            total_lines: Total number of lines
        """
        message = f"Line {line} of {total_lines}, column {column}"
        self.announce(message, priority=0)

    def announce_selection(self, text: str, char_count: int):
        """
        Announce text selection.

        Args:
            text: Selected text (truncated if long)
            char_count: Number of characters selected
        """
        if char_count == 0:
            self.announce("Selection cleared", priority=0)
        else:
            preview = text[:50] + "..." if len(text) > 50 else text
            message = f"Selected {char_count} characters: {preview}"
            self.announce(message, priority=1)

    def announce_formatting(self, format_info: Dict[str, Any]):
        """
        Announce text formatting information.

        Args:
            format_info: Dictionary with formatting details
        """
        parts = []
        if format_info.get('bold'):
            parts.append("bold")
        if format_info.get('italic'):
            parts.append("italic")
        if format_info.get('underline'):
            parts.append("underlined")

        if parts:
            message = f"Formatting: {', '.join(parts)}"
        else:
            message = "No formatting"

        self.announce(message, priority=0)

    def register_landmark(self, name: str, element_type: str, position: int):
        """
        Register a document landmark for navigation.

        Args:
            name: Landmark name
            element_type: Type (heading, table, image, etc.)
            position: Position in document
        """
        self.document_landmarks.append({
            'name': name,
            'type': element_type,
            'position': position
        })

    def get_landmarks(self, element_type: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Get document landmarks, optionally filtered by type.

        Args:
            element_type: Optional type filter

        Returns:
            List of landmarks
        """
        if element_type:
            return [lm for lm in self.document_landmarks if lm['type'] == element_type]
        return self.document_landmarks

    def clear_landmarks(self):
        """Clear all document landmarks."""
        self.document_landmarks.clear()


class KeyboardNavigation(QObject):
    """
    Enhanced keyboard navigation support.

    Features:
    - Custom keyboard shortcuts
    - Navigation between document elements
    - Keyboard-only operation mode
    - Shortcut hints
    """

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.shortcuts: Dict[str, QShortcut] = {}
        self.navigation_mode = False
        self.shortcut_hints_enabled = True

    def enable_navigation_mode(self):
        """Enable keyboard-only navigation mode."""
        self.navigation_mode = True
        self._setup_navigation_shortcuts()

    def disable_navigation_mode(self):
        """Disable keyboard-only navigation mode."""
        self.navigation_mode = False

    def _setup_navigation_shortcuts(self):
        """Setup enhanced keyboard navigation shortcuts."""
        if not self.parent():
            return

        parent = self.parent()

        # Document navigation shortcuts
        self.register_shortcut("Ctrl+Up", lambda: self._navigate_paragraph(-1),
                             "Move to previous paragraph")
        self.register_shortcut("Ctrl+Down", lambda: self._navigate_paragraph(1),
                             "Move to next paragraph")
        self.register_shortcut("Alt+Up", lambda: self._navigate_heading(-1),
                             "Move to previous heading")
        self.register_shortcut("Alt+Down", lambda: self._navigate_heading(1),
                             "Move to next heading")

        # Element navigation
        self.register_shortcut("Ctrl+Alt+T", lambda: self._navigate_to_element("table"),
                             "Move to next table")
        self.register_shortcut("Ctrl+Alt+I", lambda: self._navigate_to_element("image"),
                             "Move to next image")
        self.register_shortcut("Ctrl+Alt+L", lambda: self._navigate_to_element("list"),
                             "Move to next list")

        # UI navigation
        self.register_shortcut("F6", self._cycle_panels,
                             "Cycle through panels")
        self.register_shortcut("Shift+F6", lambda: self._cycle_panels(reverse=True),
                             "Cycle through panels (reverse)")

        # Accessibility shortcuts
        self.register_shortcut("Ctrl+Shift+A", self._toggle_accessibility_panel,
                             "Toggle accessibility panel")
        self.register_shortcut("F1", self._show_keyboard_shortcuts,
                             "Show keyboard shortcuts")

    def register_shortcut(self, key_sequence: str, callback: Callable,
                         description: str = ""):
        """
        Register a keyboard shortcut.

        Args:
            key_sequence: Keyboard shortcut (e.g., "Ctrl+S")
            callback: Function to call
            description: Description of the shortcut
        """
        if not self.parent():
            return

        shortcut = QShortcut(QKeySequence(key_sequence), self.parent())
        shortcut.activated.connect(callback)
        self.shortcuts[key_sequence] = {
            'shortcut': shortcut,
            'description': description,
            'callback': callback
        }

    def unregister_shortcut(self, key_sequence: str):
        """
        Unregister a keyboard shortcut.

        Args:
            key_sequence: Keyboard shortcut to remove
        """
        if key_sequence in self.shortcuts:
            self.shortcuts[key_sequence]['shortcut'].setEnabled(False)
            del self.shortcuts[key_sequence]

    def get_all_shortcuts(self) -> Dict[str, str]:
        """
        Get all registered shortcuts.

        Returns:
            Dictionary mapping key sequences to descriptions
        """
        return {key: info['description'] for key, info in self.shortcuts.items()}

    def _navigate_paragraph(self, direction: int):
        """Navigate to previous/next paragraph."""
        parent = self.parent()
        if hasattr(parent, 'current_editor'):
            editor = parent.current_editor()
            if editor:
                cursor = editor.textCursor()
                if direction > 0:
                    cursor.movePosition(QTextCursor.NextBlock)
                else:
                    cursor.movePosition(QTextCursor.PreviousBlock)
                editor.setTextCursor(cursor)

    def _navigate_heading(self, direction: int):
        """Navigate to previous/next heading."""
        # Implementation would search for headings in document
        pass

    def _navigate_to_element(self, element_type: str):
        """Navigate to next element of specified type."""
        # Implementation would search for specific element types
        pass

    def _cycle_panels(self, reverse: bool = False):
        """Cycle focus through UI panels."""
        parent = self.parent()
        if parent:
            # Get all focusable widgets
            focusable = [w for w in parent.findChildren(QWidget) if w.focusPolicy() != Qt.NoFocus]
            if focusable:
                current_focus = QApplication.focusWidget()
                if current_focus in focusable:
                    current_index = focusable.index(current_focus)
                    next_index = (current_index + (-1 if reverse else 1)) % len(focusable)
                    focusable[next_index].setFocus()
                elif focusable:
                    focusable[0].setFocus()

    def _toggle_accessibility_panel(self):
        """Toggle accessibility settings panel."""
        # Implementation would show/hide accessibility panel
        pass

    def _show_keyboard_shortcuts(self):
        """Show keyboard shortcuts help."""
        # Implementation would display shortcut reference
        pass


class HighContrastMode(QObject):
    """
    Provides high contrast mode support.

    Features:
    - Multiple contrast themes
    - Custom color schemes
    - Contrast ratio validation
    - Theme persistence
    """

    theme_changed = Signal(str)  # theme name

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.current_theme = "default"
        self.themes = self._define_themes()
        self.custom_themes: Dict[str, Dict[str, str]] = {}

    def _define_themes(self) -> Dict[str, Dict[str, str]]:
        """Define built-in high contrast themes."""
        return {
            "default": {
                "name": "Default",
                "background": "#FFFFFF",
                "foreground": "#000000",
                "selection_bg": "#0078D7",
                "selection_fg": "#FFFFFF"
            },
            "high_contrast_black": {
                "name": "High Contrast Black",
                "background": "#000000",
                "foreground": "#FFFFFF",
                "selection_bg": "#00FF00",
                "selection_fg": "#000000",
                "link": "#00FFFF",
                "button": "#FFFF00"
            },
            "high_contrast_white": {
                "name": "High Contrast White",
                "background": "#FFFFFF",
                "foreground": "#000000",
                "selection_bg": "#FF0000",
                "selection_fg": "#FFFFFF",
                "link": "#0000FF",
                "button": "#000080"
            },
            "high_contrast_green": {
                "name": "High Contrast Green",
                "background": "#000000",
                "foreground": "#00FF00",
                "selection_bg": "#FFFF00",
                "selection_fg": "#000000",
                "link": "#00FFFF"
            },
            "yellow_black": {
                "name": "Yellow on Black",
                "background": "#000000",
                "foreground": "#FFFF00",
                "selection_bg": "#FF00FF",
                "selection_fg": "#000000"
            }
        }

    def set_theme(self, theme_name: str) -> bool:
        """
        Set the current high contrast theme.

        Args:
            theme_name: Name of the theme to apply

        Returns:
            True if successful, False if theme not found
        """
        if theme_name not in self.themes and theme_name not in self.custom_themes:
            return False

        self.current_theme = theme_name
        self._apply_theme(theme_name)
        self.theme_changed.emit(theme_name)
        return True

    def _apply_theme(self, theme_name: str):
        """
        Apply a theme to the application.

        Args:
            theme_name: Name of the theme to apply
        """
        theme = self.themes.get(theme_name) or self.custom_themes.get(theme_name)
        if not theme:
            return

        parent = self.parent()
        if not parent:
            return

        # Create and apply palette
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(theme["background"]))
        palette.setColor(QPalette.WindowText, QColor(theme["foreground"]))
        palette.setColor(QPalette.Base, QColor(theme["background"]))
        palette.setColor(QPalette.Text, QColor(theme["foreground"]))
        palette.setColor(QPalette.Highlight, QColor(theme["selection_bg"]))
        palette.setColor(QPalette.HighlightedText, QColor(theme["selection_fg"]))

        if "button" in theme:
            palette.setColor(QPalette.Button, QColor(theme["button"]))
        if "link" in theme:
            palette.setColor(QPalette.Link, QColor(theme["link"]))

        parent.setPalette(palette)

        # Apply to all child widgets
        for widget in parent.findChildren(QWidget):
            widget.setPalette(palette)

    def add_custom_theme(self, theme_name: str, colors: Dict[str, str]):
        """
        Add a custom high contrast theme.

        Args:
            theme_name: Name for the custom theme
            colors: Dictionary of color values
        """
        self.custom_themes[theme_name] = colors

    def get_available_themes(self) -> List[str]:
        """
        Get list of available themes.

        Returns:
            List of theme names
        """
        return list(self.themes.keys()) + list(self.custom_themes.keys())

    def get_current_theme(self) -> str:
        """
        Get the current theme name.

        Returns:
            Current theme name
        """
        return self.current_theme

    def calculate_contrast_ratio(self, color1: str, color2: str) -> float:
        """
        Calculate contrast ratio between two colors.

        Args:
            color1: First color (hex format)
            color2: Second color (hex format)

        Returns:
            Contrast ratio (1-21)
        """
        def luminance(color: str) -> float:
            """Calculate relative luminance of a color."""
            rgb = QColor(color)
            r, g, b = rgb.red() / 255, rgb.green() / 255, rgb.blue() / 255

            # Convert to linear RGB
            def adjust(c):
                return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

            r, g, b = adjust(r), adjust(g), adjust(b)
            return 0.2126 * r + 0.7152 * g + 0.0722 * b

        l1 = luminance(color1)
        l2 = luminance(color2)

        lighter = max(l1, l2)
        darker = min(l1, l2)

        return (lighter + 0.05) / (darker + 0.05)

    def validate_theme_contrast(self, theme: Dict[str, str]) -> bool:
        """
        Validate that a theme meets WCAG contrast requirements.

        Args:
            theme: Theme to validate

        Returns:
            True if theme meets AA level contrast requirements
        """
        # WCAG AA requires 4.5:1 for normal text, 3:1 for large text
        bg = theme.get("background", "#FFFFFF")
        fg = theme.get("foreground", "#000000")

        ratio = self.calculate_contrast_ratio(bg, fg)
        return ratio >= 4.5


class TextToSpeech(QObject):
    """
    Provides text-to-speech functionality.

    Features:
    - Read selected text
    - Read entire document
    - Adjustable speech rate and voice
    - Pause/resume support
    """

    speech_started = Signal()
    speech_stopped = Signal()
    speech_paused = Signal()

    def __init__(self):
        super().__init__()
        self.is_speaking = False
        self.is_paused = False
        self.current_text = ""
        self.speech_rate = 1.0
        self.voice = "default"
        self._engine = None
        self._init_engine()

    def _init_engine(self):
        """Initialize the text-to-speech engine."""
        try:
            # Try to use pyttsx3 if available
            import pyttsx3
            self._engine = pyttsx3.init()
            self._engine.setProperty('rate', 150 * self.speech_rate)
        except ImportError:
            print("pyttsx3 not available, TTS will use fallback method")
            self._engine = None

    def speak(self, text: str):
        """
        Speak the given text.

        Args:
            text: Text to speak
        """
        if not text:
            return

        self.current_text = text
        self.is_speaking = True
        self.is_paused = False

        if self._engine:
            self._engine.say(text)
            self._engine.runAndWait()
        else:
            # Fallback: just emit signal
            print(f"TTS: {text[:100]}...")

        self.speech_started.emit()

    def stop(self):
        """Stop speaking."""
        if self._engine:
            self._engine.stop()

        self.is_speaking = False
        self.is_paused = False
        self.speech_stopped.emit()

    def pause(self):
        """Pause speaking."""
        if self.is_speaking and not self.is_paused:
            self.is_paused = True
            self.speech_paused.emit()

    def resume(self):
        """Resume speaking."""
        if self.is_paused:
            self.is_paused = False
            # Resume from where we left off
            if self._engine and self.current_text:
                self.speak(self.current_text)

    def set_rate(self, rate: float):
        """
        Set speech rate.

        Args:
            rate: Speech rate (0.5 - 2.0, default 1.0)
        """
        self.speech_rate = max(0.5, min(2.0, rate))
        if self._engine:
            self._engine.setProperty('rate', 150 * self.speech_rate)

    def set_voice(self, voice_id: str):
        """
        Set voice.

        Args:
            voice_id: Voice identifier
        """
        self.voice = voice_id
        if self._engine:
            voices = self._engine.getProperty('voices')
            for voice in voices:
                if voice_id in voice.id:
                    self._engine.setProperty('voice', voice.id)
                    break

    def get_available_voices(self) -> List[str]:
        """
        Get list of available voices.

        Returns:
            List of voice identifiers
        """
        if self._engine:
            voices = self._engine.getProperty('voices')
            return [voice.id for voice in voices]
        return ["default"]

    def read_selection(self, text_edit: QTextEdit):
        """
        Read the selected text from a text editor.

        Args:
            text_edit: QTextEdit widget
        """
        cursor = text_edit.textCursor()
        if cursor.hasSelection():
            text = cursor.selectedText()
            self.speak(text)

    def read_document(self, text_edit: QTextEdit):
        """
        Read the entire document.

        Args:
            text_edit: QTextEdit widget
        """
        text = text_edit.toPlainText()
        self.speak(text)


class AccessibilityManager(QObject):
    """
    Central manager for all accessibility features.

    Coordinates:
    - Screen reader support
    - Keyboard navigation
    - High contrast mode
    - Text-to-speech
    """

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.screen_reader = ScreenReaderSupport(parent)
        self.keyboard_nav = KeyboardNavigation(parent)
        self.high_contrast = HighContrastMode(parent)
        self.tts = TextToSpeech()

        self.accessibility_level = AccessibilityLevel.BASIC
        self.features_enabled: Dict[str, bool] = {
            'screen_reader': True,
            'keyboard_nav': True,
            'high_contrast': False,
            'tts': True
        }

    def set_accessibility_level(self, level: AccessibilityLevel):
        """
        Set the overall accessibility level.

        Args:
            level: Accessibility level to set
        """
        self.accessibility_level = level

        if level == AccessibilityLevel.NONE:
            self.disable_all_features()
        elif level == AccessibilityLevel.BASIC:
            self.enable_feature('screen_reader')
            self.enable_feature('keyboard_nav')
        elif level == AccessibilityLevel.FULL:
            self.enable_all_features()

    def enable_feature(self, feature_name: str):
        """
        Enable a specific accessibility feature.

        Args:
            feature_name: Name of the feature to enable
        """
        if feature_name not in self.features_enabled:
            return

        self.features_enabled[feature_name] = True

        if feature_name == 'screen_reader':
            self.screen_reader.enable()
        elif feature_name == 'keyboard_nav':
            self.keyboard_nav.enable_navigation_mode()

    def disable_feature(self, feature_name: str):
        """
        Disable a specific accessibility feature.

        Args:
            feature_name: Name of the feature to disable
        """
        if feature_name not in self.features_enabled:
            return

        self.features_enabled[feature_name] = False

        if feature_name == 'screen_reader':
            self.screen_reader.disable()
        elif feature_name == 'keyboard_nav':
            self.keyboard_nav.disable_navigation_mode()
        elif feature_name == 'tts':
            self.tts.stop()

    def enable_all_features(self):
        """Enable all accessibility features."""
        for feature in self.features_enabled.keys():
            self.enable_feature(feature)

    def disable_all_features(self):
        """Disable all accessibility features."""
        for feature in self.features_enabled.keys():
            self.disable_feature(feature)

    def is_feature_enabled(self, feature_name: str) -> bool:
        """
        Check if a feature is enabled.

        Args:
            feature_name: Name of the feature

        Returns:
            True if enabled, False otherwise
        """
        return self.features_enabled.get(feature_name, False)

    def get_feature_status(self) -> Dict[str, bool]:
        """
        Get status of all features.

        Returns:
            Dictionary of feature statuses
        """
        return self.features_enabled.copy()
