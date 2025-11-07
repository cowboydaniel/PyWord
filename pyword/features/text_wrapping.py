"""
Text Wrapping Around Objects for PyWord.

This module provides functionality for text wrapping around images, shapes,
and other objects in the document.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
                               QPushButton, QGroupBox, QDoubleSpinBox, QRadioButton,
                               QButtonGroup, QFormLayout, QSlider, QCheckBox)
from PySide6.QtCore import Qt, QRectF, QPointF
from PySide6.QtGui import (QTextFrameFormat, QTextImageFormat, QTextCursor,
                          QTextFrame, QTextLength)


class WrapType:
    """Text wrapping types."""
    INLINE = "In Line with Text"
    SQUARE = "Square"
    TIGHT = "Tight"
    THROUGH = "Through"
    TOP_BOTTOM = "Top and Bottom"
    BEHIND_TEXT = "Behind Text"
    IN_FRONT = "In Front of Text"


class WrapSide:
    """Which side text wraps around."""
    BOTH = "Both Sides"
    LEFT = "Left Only"
    RIGHT = "Right Only"
    LARGEST = "Largest Only"


class TextWrappingManager:
    """Manages text wrapping around objects."""

    def __init__(self, parent):
        self.parent = parent
        self.wrapped_objects = {}  # Dictionary to store wrapping info for objects

    def set_image_wrapping(self, image_format, wrap_type=WrapType.SQUARE,
                          wrap_side=WrapSide.BOTH, distance=10.0):
        """Set text wrapping for an image."""
        if not isinstance(image_format, QTextImageFormat):
            return False

        # Store wrapping information
        image_name = image_format.name()
        self.wrapped_objects[image_name] = {
            'type': 'image',
            'wrap_type': wrap_type,
            'wrap_side': wrap_side,
            'distance': distance
        }

        # Apply wrapping using QTextImageFormat properties
        if wrap_type == WrapType.INLINE:
            # Default inline positioning
            pass

        elif wrap_type == WrapType.SQUARE:
            # Set the image to float
            image_format.setProperty(QTextFrameFormat.Property.FloatRight, True)

        elif wrap_type == WrapType.TOP_BOTTOM:
            # Image on its own line
            image_format.setProperty(QTextFrameFormat.Property.ClearLeft, True)
            image_format.setProperty(QTextFrameFormat.Property.ClearRight, True)

        elif wrap_type == WrapType.BEHIND_TEXT:
            # Set lower z-index (simulate background)
            image_format.setProperty(QTextFormat.Property.BackgroundBrush, Qt.GlobalColor.transparent)

        elif wrap_type == WrapType.IN_FRONT:
            # Set higher z-index (simulate foreground)
            pass

        return True

    def set_frame_wrapping(self, frame, wrap_type=WrapType.SQUARE,
                          wrap_side=WrapSide.BOTH, distance=10.0):
        """Set text wrapping for a frame/shape."""
        if not isinstance(frame, QTextFrame):
            return False

        frame_format = frame.frameFormat()

        # Store wrapping information
        frame_id = id(frame)
        self.wrapped_objects[frame_id] = {
            'type': 'frame',
            'wrap_type': wrap_type,
            'wrap_side': wrap_side,
            'distance': distance
        }

        # Apply wrapping styles
        if wrap_type == WrapType.INLINE:
            frame_format.setPosition(QTextFrameFormat.Position.InFlow)

        elif wrap_type == WrapType.SQUARE or wrap_type == WrapType.TIGHT:
            frame_format.setPosition(QTextFrameFormat.Position.FloatRight)
            frame_format.setMargin(distance)

        elif wrap_type == WrapType.TOP_BOTTOM:
            frame_format.setPosition(QTextFrameFormat.Position.InFlow)
            frame_format.setPageBreakPolicy(
                QTextFrameFormat.PageBreakFlag.PageBreak_AlwaysBefore |
                QTextFrameFormat.PageBreakFlag.PageBreak_AlwaysAfter
            )

        frame.setFrameFormat(frame_format)
        return True

    def set_table_wrapping(self, table, wrap_type=WrapType.SQUARE, distance=10.0):
        """Set text wrapping for a table."""
        table_format = table.format()

        # Apply wrapping
        if wrap_type == WrapType.INLINE:
            table_format.setPosition(QTextFrameFormat.Position.InFlow)

        elif wrap_type == WrapType.SQUARE:
            table_format.setPosition(QTextFrameFormat.Position.FloatRight)
            table_format.setMargin(distance)

        elif wrap_type == WrapType.TOP_BOTTOM:
            table_format.setPosition(QTextFrameFormat.Position.InFlow)

        table.setFormat(table_format)
        return True

    def get_current_object_wrapping(self):
        """Get wrapping information for the object at cursor position."""
        cursor = self.parent.textCursor()

        # Check for image
        char_format = cursor.charFormat()
        if char_format.isImageFormat():
            image_format = char_format.toImageFormat()
            image_name = image_format.name()
            if image_name in self.wrapped_objects:
                return self.wrapped_objects[image_name]

        # Check for frame
        frame = cursor.currentFrame()
        if frame:
            frame_id = id(frame)
            if frame_id in self.wrapped_objects:
                return self.wrapped_objects[frame_id]

        # Check for table
        table = cursor.currentTable()
        if table:
            table_id = id(table)
            if table_id in self.wrapped_objects:
                return self.wrapped_objects[table_id]

        return None

    def apply_wrapping_to_current_object(self, wrap_type, wrap_side=WrapSide.BOTH, distance=10.0):
        """Apply text wrapping to the currently selected object."""
        cursor = self.parent.textCursor()

        # Try image first
        char_format = cursor.charFormat()
        if char_format.isImageFormat():
            image_format = char_format.toImageFormat()
            return self.set_image_wrapping(image_format, wrap_type, wrap_side, distance)

        # Try frame
        frame = cursor.currentFrame()
        if frame and frame != self.parent.document().rootFrame():
            return self.set_frame_wrapping(frame, wrap_type, wrap_side, distance)

        # Try table
        table = cursor.currentTable()
        if table:
            return self.set_table_wrapping(table, wrap_type, distance)

        return False

    def set_object_position(self, position_type="inline", horizontal_pos=0, vertical_pos=0):
        """Set the position of an object."""
        cursor = self.parent.textCursor()

        # For frames
        frame = cursor.currentFrame()
        if frame and frame != self.parent.document().rootFrame():
            frame_format = frame.frameFormat()

            if position_type == "inline":
                frame_format.setPosition(QTextFrameFormat.Position.InFlow)
            elif position_type == "absolute":
                # Note: QTextEdit doesn't fully support absolute positioning
                # This is a simplified implementation
                frame_format.setPosition(QTextFrameFormat.Position.FloatRight)
                frame_format.setLeftMargin(horizontal_pos)
                frame_format.setTopMargin(vertical_pos)

            frame.setFrameFormat(frame_format)
            return True

        return False

    def bring_to_front(self):
        """Bring the current object to the front."""
        # This is a simplified implementation
        # Full z-order management would require more complex handling
        cursor = self.parent.textCursor()
        return self.apply_wrapping_to_current_object(WrapType.IN_FRONT)

    def send_to_back(self):
        """Send the current object to the back."""
        cursor = self.parent.textCursor()
        return self.apply_wrapping_to_current_object(WrapType.BEHIND_TEXT)


class TextWrappingDialog(QDialog):
    """Dialog for configuring text wrapping around objects."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Text Wrapping")
        self.setModal(True)
        self.setMinimumWidth(450)

        self.setup_ui()
        self.load_current_settings()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Wrapping style
        style_group = QGroupBox("Wrapping Style")
        style_layout = QVBoxLayout()

        self.wrap_button_group = QButtonGroup()

        self.inline_radio = QRadioButton("In Line with Text")
        self.inline_radio.setToolTip("Object is positioned inline with text")

        self.square_radio = QRadioButton("Square")
        self.square_radio.setToolTip("Text wraps in a square around the object")

        self.tight_radio = QRadioButton("Tight")
        self.tight_radio.setToolTip("Text wraps tightly around the object")

        self.through_radio = QRadioButton("Through")
        self.through_radio.setToolTip("Text flows through transparent areas")

        self.top_bottom_radio = QRadioButton("Top and Bottom")
        self.top_bottom_radio.setToolTip("Text appears above and below the object")

        self.behind_radio = QRadioButton("Behind Text")
        self.behind_radio.setToolTip("Object appears behind the text")

        self.front_radio = QRadioButton("In Front of Text")
        self.front_radio.setToolTip("Object appears in front of the text")

        self.wrap_button_group.addButton(self.inline_radio)
        self.wrap_button_group.addButton(self.square_radio)
        self.wrap_button_group.addButton(self.tight_radio)
        self.wrap_button_group.addButton(self.through_radio)
        self.wrap_button_group.addButton(self.top_bottom_radio)
        self.wrap_button_group.addButton(self.behind_radio)
        self.wrap_button_group.addButton(self.front_radio)

        self.square_radio.setChecked(True)

        style_layout.addWidget(self.inline_radio)
        style_layout.addWidget(self.square_radio)
        style_layout.addWidget(self.tight_radio)
        style_layout.addWidget(self.through_radio)
        style_layout.addWidget(self.top_bottom_radio)
        style_layout.addWidget(self.behind_radio)
        style_layout.addWidget(self.front_radio)

        style_group.setLayout(style_layout)
        layout.addWidget(style_group)

        # Wrap side
        side_group = QGroupBox("Wrap Text")
        side_layout = QVBoxLayout()

        self.side_button_group = QButtonGroup()

        self.both_sides_radio = QRadioButton("Both Sides")
        self.left_only_radio = QRadioButton("Left Only")
        self.right_only_radio = QRadioButton("Right Only")
        self.largest_radio = QRadioButton("Largest Side Only")

        self.side_button_group.addButton(self.both_sides_radio)
        self.side_button_group.addButton(self.left_only_radio)
        self.side_button_group.addButton(self.right_only_radio)
        self.side_button_group.addButton(self.largest_radio)

        self.both_sides_radio.setChecked(True)

        side_layout.addWidget(self.both_sides_radio)
        side_layout.addWidget(self.left_only_radio)
        side_layout.addWidget(self.right_only_radio)
        side_layout.addWidget(self.largest_radio)

        side_group.setLayout(side_layout)
        layout.addWidget(side_group)

        # Distance from text
        distance_group = QGroupBox("Distance from Text")
        distance_layout = QFormLayout()

        self.distance_spin = QDoubleSpinBox()
        self.distance_spin.setMinimum(0)
        self.distance_spin.setMaximum(100)
        self.distance_spin.setValue(10)
        self.distance_spin.setSuffix(" pt")

        distance_layout.addRow("Distance:", self.distance_spin)

        distance_group.setLayout(distance_layout)
        layout.addWidget(distance_group)

        # Position options
        position_group = QGroupBox("Position")
        position_layout = QHBoxLayout()

        bring_front_button = QPushButton("Bring to Front")
        bring_front_button.clicked.connect(self.bring_to_front)

        send_back_button = QPushButton("Send to Back")
        send_back_button.clicked.connect(self.send_to_back)

        position_layout.addWidget(bring_front_button)
        position_layout.addWidget(send_back_button)

        position_group.setLayout(position_layout)
        layout.addWidget(position_group)

        # Buttons
        button_layout = QHBoxLayout()
        apply_button = QPushButton("Apply")
        apply_button.clicked.connect(self.apply_wrapping)
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(apply_button)
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def load_current_settings(self):
        """Load current wrapping settings for the selected object."""
        current_wrapping = self.manager.get_current_object_wrapping()
        if current_wrapping:
            wrap_type = current_wrapping.get('wrap_type', WrapType.SQUARE)
            wrap_side = current_wrapping.get('wrap_side', WrapSide.BOTH)
            distance = current_wrapping.get('distance', 10.0)

            # Set wrap type radio button
            if wrap_type == WrapType.INLINE:
                self.inline_radio.setChecked(True)
            elif wrap_type == WrapType.SQUARE:
                self.square_radio.setChecked(True)
            elif wrap_type == WrapType.TIGHT:
                self.tight_radio.setChecked(True)
            elif wrap_type == WrapType.THROUGH:
                self.through_radio.setChecked(True)
            elif wrap_type == WrapType.TOP_BOTTOM:
                self.top_bottom_radio.setChecked(True)
            elif wrap_type == WrapType.BEHIND_TEXT:
                self.behind_radio.setChecked(True)
            elif wrap_type == WrapType.IN_FRONT:
                self.front_radio.setChecked(True)

            # Set wrap side radio button
            if wrap_side == WrapSide.BOTH:
                self.both_sides_radio.setChecked(True)
            elif wrap_side == WrapSide.LEFT:
                self.left_only_radio.setChecked(True)
            elif wrap_side == WrapSide.RIGHT:
                self.right_only_radio.setChecked(True)
            elif wrap_side == WrapSide.LARGEST:
                self.largest_radio.setChecked(True)

            # Set distance
            self.distance_spin.setValue(distance)

    def get_selected_wrap_type(self):
        """Get the selected wrap type."""
        if self.inline_radio.isChecked():
            return WrapType.INLINE
        elif self.square_radio.isChecked():
            return WrapType.SQUARE
        elif self.tight_radio.isChecked():
            return WrapType.TIGHT
        elif self.through_radio.isChecked():
            return WrapType.THROUGH
        elif self.top_bottom_radio.isChecked():
            return WrapType.TOP_BOTTOM
        elif self.behind_radio.isChecked():
            return WrapType.BEHIND_TEXT
        elif self.front_radio.isChecked():
            return WrapType.IN_FRONT
        return WrapType.SQUARE

    def get_selected_wrap_side(self):
        """Get the selected wrap side."""
        if self.both_sides_radio.isChecked():
            return WrapSide.BOTH
        elif self.left_only_radio.isChecked():
            return WrapSide.LEFT
        elif self.right_only_radio.isChecked():
            return WrapSide.RIGHT
        elif self.largest_radio.isChecked():
            return WrapSide.LARGEST
        return WrapSide.BOTH

    def apply_wrapping(self):
        """Apply the selected wrapping settings."""
        wrap_type = self.get_selected_wrap_type()
        wrap_side = self.get_selected_wrap_side()
        distance = self.distance_spin.value()

        self.manager.apply_wrapping_to_current_object(wrap_type, wrap_side, distance)

    def bring_to_front(self):
        """Bring object to front."""
        self.manager.bring_to_front()

    def send_to_back(self):
        """Send object to back."""
        self.manager.send_to_back()

    def accept(self):
        """Handle OK button click."""
        self.apply_wrapping()
        super().accept()
