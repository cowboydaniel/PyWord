"""
Page Backgrounds and Watermarks for PyWord.

This module provides functionality for adding backgrounds and watermarks to document pages.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
                               QPushButton, QGroupBox, QLineEdit, QSpinBox, QSlider,
                               QColorDialog, QFileDialog, QRadioButton, QButtonGroup,
                               QCheckBox, QTabWidget, QWidget, QFormLayout)
from PySide6.QtCore import Qt, QRectF, QPointF
from PySide6.QtGui import (QColor, QBrush, QPainter, QPixmap, QImage, QFont,
                          QTextFrameFormat, QPen, QLinearGradient, QRadialGradient)


class WatermarkType:
    """Enumeration of watermark types."""
    TEXT = "Text"
    IMAGE = "Image"
    NONE = "None"


class BackgroundType:
    """Enumeration of background types."""
    SOLID = "Solid Color"
    GRADIENT = "Gradient"
    IMAGE = "Image"
    NONE = "None"


class PageBackground:
    """Represents a page background."""

    def __init__(self):
        self.type = BackgroundType.NONE
        self.color = QColor(Qt.GlobalColor.white)
        self.gradient_start = QColor(Qt.GlobalColor.white)
        self.gradient_end = QColor(Qt.GlobalColor.lightGray)
        self.gradient_type = "Linear"  # Linear or Radial
        self.image_path = ""
        self.image_opacity = 100

    def apply_to_document(self, document):
        """Apply background to a document."""
        # Set document background
        if self.type == BackgroundType.SOLID:
            document.setProperty("background_color", self.color)
        elif self.type == BackgroundType.GRADIENT:
            document.setProperty("gradient_start", self.gradient_start)
            document.setProperty("gradient_end", self.gradient_end)
            document.setProperty("gradient_type", self.gradient_type)
        elif self.type == BackgroundType.IMAGE:
            document.setProperty("background_image", self.image_path)
            document.setProperty("background_opacity", self.image_opacity)


class Watermark:
    """Represents a watermark."""

    def __init__(self):
        self.type = WatermarkType.NONE
        self.text = "CONFIDENTIAL"
        self.font = QFont("Arial", 48)
        self.color = QColor(200, 200, 200, 128)  # Semi-transparent gray
        self.angle = 45  # Diagonal
        self.image_path = ""
        self.opacity = 50  # 0-100
        self.scale = 100  # Percentage

    def render(self, painter, page_rect):
        """Render the watermark on a page."""
        if self.type == WatermarkType.NONE:
            return

        painter.save()

        if self.type == WatermarkType.TEXT:
            # Calculate center position
            center = page_rect.center()

            # Set up painter
            painter.setOpacity(self.opacity / 100.0)
            painter.translate(center)
            painter.rotate(self.angle)

            # Set font and color
            painter.setFont(self.font)
            painter.setPen(self.color)

            # Draw text centered
            text_rect = QRectF(-500, -100, 1000, 200)
            painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, self.text)

        elif self.type == WatermarkType.IMAGE:
            # Load and draw image
            if self.image_path:
                image = QImage(self.image_path)
                if not image.isNull():
                    # Calculate scaled size
                    scaled_width = page_rect.width() * self.scale / 100.0
                    scaled_height = image.height() * scaled_width / image.width()

                    # Center the image
                    x = (page_rect.width() - scaled_width) / 2
                    y = (page_rect.height() - scaled_height) / 2

                    target_rect = QRectF(x, y, scaled_width, scaled_height)

                    painter.setOpacity(self.opacity / 100.0)
                    painter.drawImage(target_rect, image)

        painter.restore()


class PageBackgroundManager:
    """Manages page backgrounds for a document."""

    def __init__(self, parent):
        self.parent = parent
        self.background = PageBackground()
        self.watermark = Watermark()

    def set_background_color(self, color):
        """Set a solid background color."""
        self.background.type = BackgroundType.SOLID
        self.background.color = color
        self.apply_background()

    def set_background_gradient(self, start_color, end_color, gradient_type="Linear"):
        """Set a gradient background."""
        self.background.type = BackgroundType.GRADIENT
        self.background.gradient_start = start_color
        self.background.gradient_end = end_color
        self.background.gradient_type = gradient_type
        self.apply_background()

    def set_background_image(self, image_path, opacity=100):
        """Set an image background."""
        self.background.type = BackgroundType.IMAGE
        self.background.image_path = image_path
        self.background.image_opacity = opacity
        self.apply_background()

    def remove_background(self):
        """Remove the page background."""
        self.background.type = BackgroundType.NONE
        self.apply_background()

    def apply_background(self):
        """Apply the current background to the document."""
        if hasattr(self.parent, 'document'):
            doc = self.parent.document()
            self.background.apply_to_document(doc)
            self.parent.viewport().update()

    def set_text_watermark(self, text, font=None, color=None, angle=45, opacity=50):
        """Set a text watermark."""
        self.watermark.type = WatermarkType.TEXT
        self.watermark.text = text
        if font:
            self.watermark.font = font
        if color:
            self.watermark.color = color
        self.watermark.angle = angle
        self.watermark.opacity = opacity
        self.apply_watermark()

    def set_image_watermark(self, image_path, opacity=50, scale=100):
        """Set an image watermark."""
        self.watermark.type = WatermarkType.IMAGE
        self.watermark.image_path = image_path
        self.watermark.opacity = opacity
        self.watermark.scale = scale
        self.apply_watermark()

    def remove_watermark(self):
        """Remove the watermark."""
        self.watermark.type = WatermarkType.NONE
        self.apply_watermark()

    def apply_watermark(self):
        """Apply the current watermark to the document."""
        if hasattr(self.parent, 'viewport'):
            self.parent.viewport().update()


class PageBackgroundDialog(QDialog):
    """Dialog for configuring page backgrounds and watermarks."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Page Background and Watermark")
        self.setModal(True)
        self.setMinimumWidth(500)

        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Tab widget for background and watermark
        tabs = QTabWidget()

        # Background tab
        background_tab = self.create_background_tab()
        tabs.addTab(background_tab, "Background")

        # Watermark tab
        watermark_tab = self.create_watermark_tab()
        tabs.addTab(watermark_tab, "Watermark")

        layout.addWidget(tabs)

        # Buttons
        button_layout = QHBoxLayout()
        apply_button = QPushButton("Apply")
        apply_button.clicked.connect(self.apply_settings)
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

    def create_background_tab(self):
        """Create the background configuration tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Background type
        type_group = QGroupBox("Background Type")
        type_layout = QVBoxLayout()

        self.bg_button_group = QButtonGroup()
        self.bg_none_radio = QRadioButton("None")
        self.bg_solid_radio = QRadioButton("Solid Color")
        self.bg_gradient_radio = QRadioButton("Gradient")
        self.bg_image_radio = QRadioButton("Image")

        self.bg_button_group.addButton(self.bg_none_radio)
        self.bg_button_group.addButton(self.bg_solid_radio)
        self.bg_button_group.addButton(self.bg_gradient_radio)
        self.bg_button_group.addButton(self.bg_image_radio)

        self.bg_none_radio.setChecked(True)

        type_layout.addWidget(self.bg_none_radio)
        type_layout.addWidget(self.bg_solid_radio)
        type_layout.addWidget(self.bg_gradient_radio)
        type_layout.addWidget(self.bg_image_radio)

        type_group.setLayout(type_layout)
        layout.addWidget(type_group)

        # Solid color options
        self.solid_group = QGroupBox("Solid Color")
        solid_layout = QHBoxLayout()
        self.bg_color_button = QPushButton("Choose Color")
        self.bg_color_button.clicked.connect(self.choose_background_color)
        solid_layout.addWidget(self.bg_color_button)
        self.solid_group.setLayout(solid_layout)
        layout.addWidget(self.solid_group)

        # Gradient options
        self.gradient_group = QGroupBox("Gradient")
        gradient_layout = QFormLayout()
        self.gradient_start_button = QPushButton("Start Color")
        self.gradient_start_button.clicked.connect(self.choose_gradient_start)
        self.gradient_end_button = QPushButton("End Color")
        self.gradient_end_button.clicked.connect(self.choose_gradient_end)
        self.gradient_type_combo = QComboBox()
        self.gradient_type_combo.addItems(["Linear", "Radial"])
        gradient_layout.addRow("Start Color:", self.gradient_start_button)
        gradient_layout.addRow("End Color:", self.gradient_end_button)
        gradient_layout.addRow("Type:", self.gradient_type_combo)
        self.gradient_group.setLayout(gradient_layout)
        layout.addWidget(self.gradient_group)

        # Image options
        self.image_group = QGroupBox("Image")
        image_layout = QVBoxLayout()
        self.bg_image_button = QPushButton("Choose Image")
        self.bg_image_button.clicked.connect(self.choose_background_image)
        image_layout.addWidget(self.bg_image_button)
        self.bg_image_path_label = QLabel("No image selected")
        image_layout.addWidget(self.bg_image_path_label)
        self.image_group.setLayout(image_layout)
        layout.addWidget(self.image_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_watermark_tab(self):
        """Create the watermark configuration tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Watermark type
        type_group = QGroupBox("Watermark Type")
        type_layout = QVBoxLayout()

        self.wm_button_group = QButtonGroup()
        self.wm_none_radio = QRadioButton("None")
        self.wm_text_radio = QRadioButton("Text")
        self.wm_image_radio = QRadioButton("Image")

        self.wm_button_group.addButton(self.wm_none_radio)
        self.wm_button_group.addButton(self.wm_text_radio)
        self.wm_button_group.addButton(self.wm_image_radio)

        self.wm_none_radio.setChecked(True)

        type_layout.addWidget(self.wm_none_radio)
        type_layout.addWidget(self.wm_text_radio)
        type_layout.addWidget(self.wm_image_radio)

        type_group.setLayout(type_layout)
        layout.addWidget(type_group)

        # Text watermark options
        self.text_wm_group = QGroupBox("Text Watermark")
        text_layout = QFormLayout()
        self.wm_text_edit = QLineEdit("CONFIDENTIAL")
        self.wm_angle_spin = QSpinBox()
        self.wm_angle_spin.setRange(0, 360)
        self.wm_angle_spin.setValue(45)
        self.wm_opacity_slider = QSlider(Qt.Orientation.Horizontal)
        self.wm_opacity_slider.setRange(0, 100)
        self.wm_opacity_slider.setValue(50)
        text_layout.addRow("Text:", self.wm_text_edit)
        text_layout.addRow("Angle:", self.wm_angle_spin)
        text_layout.addRow("Opacity:", self.wm_opacity_slider)
        self.text_wm_group.setLayout(text_layout)
        layout.addWidget(self.text_wm_group)

        # Image watermark options
        self.image_wm_group = QGroupBox("Image Watermark")
        image_layout = QVBoxLayout()
        self.wm_image_button = QPushButton("Choose Image")
        self.wm_image_button.clicked.connect(self.choose_watermark_image)
        image_layout.addWidget(self.wm_image_button)
        self.wm_image_path_label = QLabel("No image selected")
        image_layout.addWidget(self.wm_image_path_label)
        self.image_wm_group.setLayout(image_layout)
        layout.addWidget(self.image_wm_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def choose_background_color(self):
        """Choose background color."""
        color = QColorDialog.getColor()
        if color.isValid():
            self.bg_color = color

    def choose_gradient_start(self):
        """Choose gradient start color."""
        color = QColorDialog.getColor()
        if color.isValid():
            self.gradient_start_color = color

    def choose_gradient_end(self):
        """Choose gradient end color."""
        color = QColorDialog.getColor()
        if color.isValid():
            self.gradient_end_color = color

    def choose_background_image(self):
        """Choose background image."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Choose Background Image", "",
            "Images (*.png *.jpg *.jpeg *.bmp *.gif)"
        )
        if file_path:
            self.bg_image_path = file_path
            self.bg_image_path_label.setText(file_path)

    def choose_watermark_image(self):
        """Choose watermark image."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Choose Watermark Image", "",
            "Images (*.png *.jpg *.jpeg *.bmp *.gif)"
        )
        if file_path:
            self.wm_image_path = file_path
            self.wm_image_path_label.setText(file_path)

    def apply_settings(self):
        """Apply the current settings."""
        # Apply background
        if self.bg_solid_radio.isChecked():
            if hasattr(self, 'bg_color'):
                self.manager.set_background_color(self.bg_color)
        elif self.bg_gradient_radio.isChecked():
            if hasattr(self, 'gradient_start_color') and hasattr(self, 'gradient_end_color'):
                self.manager.set_background_gradient(
                    self.gradient_start_color,
                    self.gradient_end_color,
                    self.gradient_type_combo.currentText()
                )
        elif self.bg_image_radio.isChecked():
            if hasattr(self, 'bg_image_path'):
                self.manager.set_background_image(self.bg_image_path)
        else:
            self.manager.remove_background()

        # Apply watermark
        if self.wm_text_radio.isChecked():
            self.manager.set_text_watermark(
                self.wm_text_edit.text(),
                angle=self.wm_angle_spin.value(),
                opacity=self.wm_opacity_slider.value()
            )
        elif self.wm_image_radio.isChecked():
            if hasattr(self, 'wm_image_path'):
                self.manager.set_image_watermark(self.wm_image_path)
        else:
            self.manager.remove_watermark()

    def accept(self):
        """Handle OK button click."""
        self.apply_settings()
        super().accept()
