"""
Advanced Image Handling for PyWord.

This module provides advanced image manipulation features including:
- Image resizing and cropping
- Image rotation and flipping
- Image effects (brightness, contrast, filters)
- Image compression
- Image positioning and layout
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
                               QPushButton, QGroupBox, QSpinBox, QDoubleSpinBox,
                               QSlider, QFormLayout, QCheckBox, QTabWidget, QWidget,
                               QFileDialog, QMessageBox)
from PySide6.QtCore import Qt, QSize, QRectF
from PySide6.QtGui import (QImage, QPixmap, QTransform, QTextImageFormat, QTextCursor,
                          QImageReader, QImageWriter, QPainter, QColor, QTextDocument)
import os


class ImageEffect:
    """Image effect types."""
    NONE = "None"
    GRAYSCALE = "Grayscale"
    SEPIA = "Sepia"
    INVERT = "Invert"
    BLUR = "Blur"
    SHARPEN = "Sharpen"


class AdvancedImageManager:
    """Manages advanced image operations."""

    def __init__(self, parent):
        self.parent = parent
        self.images = {}  # Store image information

    def get_current_image(self):
        """Get the image at the current cursor position."""
        cursor = self.parent.textCursor()
        char_format = cursor.charFormat()

        if char_format.isImageFormat():
            image_format = char_format.toImageFormat()
            return image_format

        return None

    def insert_image(self, image_path, width=None, height=None):
        """Insert an image into the document."""
        if not os.path.exists(image_path):
            return False

        cursor = self.parent.textCursor()

        # Load the image
        image = QImage(image_path)
        if image.isNull():
            return False

        # Create image format
        image_format = QTextImageFormat()
        image_format.setName(image_path)

        # Set size
        if width and height:
            image_format.setWidth(width)
            image_format.setHeight(height)
        elif width:
            # Maintain aspect ratio
            aspect_ratio = image.height() / image.width()
            image_format.setWidth(width)
            image_format.setHeight(width * aspect_ratio)
        elif height:
            # Maintain aspect ratio
            aspect_ratio = image.width() / image.height()
            image_format.setHeight(height)
            image_format.setWidth(height * aspect_ratio)
        else:
            # Use original size
            image_format.setWidth(image.width())
            image_format.setHeight(image.height())

        # Add image to document resources
        self.parent.document().addResource(QTextDocument.ResourceType.ImageResource,
                                          image_path, image)

        # Insert the image
        cursor.insertImage(image_format)

        # Store image info
        self.images[image_path] = {
            'original_size': (image.width(), image.height()),
            'current_size': (image_format.width(), image_format.height()),
            'path': image_path
        }

        return True

    def resize_image(self, width=None, height=None, maintain_aspect_ratio=True):
        """Resize the current image."""
        image_format = self.get_current_image()
        if not image_format:
            return False

        cursor = self.parent.textCursor()
        image_path = image_format.name()

        # Load original image
        image = QImage(image_path)
        if image.isNull():
            return False

        # Calculate new size
        if maintain_aspect_ratio:
            aspect_ratio = image.height() / image.width()
            if width and not height:
                height = width * aspect_ratio
            elif height and not width:
                width = height / aspect_ratio
            elif width and height:
                # Use width and calculate height
                height = width * aspect_ratio

        # Update format
        new_format = QTextImageFormat()
        new_format.setName(image_path)
        if width:
            new_format.setWidth(width)
        if height:
            new_format.setHeight(height)

        # Apply the new format
        cursor.setCharFormat(new_format)

        return True

    def rotate_image(self, angle):
        """Rotate the current image."""
        image_format = self.get_current_image()
        if not image_format:
            return False

        image_path = image_format.name()
        image = QImage(image_path)

        if image.isNull():
            return False

        # Rotate the image
        transform = QTransform().rotate(angle)
        rotated_image = image.transformed(transform, Qt.TransformationMode.SmoothTransformation)

        # Save rotated image temporarily
        temp_path = f"{image_path}_rotated_{angle}.png"
        rotated_image.save(temp_path)

        # Update document resource
        self.parent.document().addResource(QTextDocument.ResourceType.ImageResource,
                                          temp_path, rotated_image)

        # Update image format
        cursor = self.parent.textCursor()
        new_format = QTextImageFormat()
        new_format.setName(temp_path)
        new_format.setWidth(rotated_image.width())
        new_format.setHeight(rotated_image.height())

        cursor.setCharFormat(new_format)

        return True

    def flip_image(self, horizontal=True, vertical=False):
        """Flip the current image."""
        image_format = self.get_current_image()
        if not image_format:
            return False

        image_path = image_format.name()
        image = QImage(image_path)

        if image.isNull():
            return False

        # Flip the image
        flipped_image = image.mirrored(horizontal, vertical)

        # Save flipped image temporarily
        flip_type = "h" if horizontal else "v"
        temp_path = f"{image_path}_flipped_{flip_type}.png"
        flipped_image.save(temp_path)

        # Update document resource
        self.parent.document().addResource(QTextDocument.ResourceType.ImageResource,
                                          temp_path, flipped_image)

        # Update image format
        cursor = self.parent.textCursor()
        new_format = QTextImageFormat()
        new_format.setName(temp_path)
        new_format.setWidth(flipped_image.width())
        new_format.setHeight(flipped_image.height())

        cursor.setCharFormat(new_format)

        return True

    def crop_image(self, left, top, width, height):
        """Crop the current image."""
        image_format = self.get_current_image()
        if not image_format:
            return False

        image_path = image_format.name()
        image = QImage(image_path)

        if image.isNull():
            return False

        # Crop the image
        cropped_image = image.copy(left, top, width, height)

        # Save cropped image temporarily
        temp_path = f"{image_path}_cropped.png"
        cropped_image.save(temp_path)

        # Update document resource
        self.parent.document().addResource(QTextDocument.ResourceType.ImageResource,
                                          temp_path, cropped_image)

        # Update image format
        cursor = self.parent.textCursor()
        new_format = QTextImageFormat()
        new_format.setName(temp_path)
        new_format.setWidth(cropped_image.width())
        new_format.setHeight(cropped_image.height())

        cursor.setCharFormat(new_format)

        return True

    def apply_effect(self, effect_type):
        """Apply an effect to the current image."""
        image_format = self.get_current_image()
        if not image_format:
            return False

        image_path = image_format.name()
        image = QImage(image_path)

        if image.isNull():
            return False

        # Apply effect
        if effect_type == ImageEffect.GRAYSCALE:
            image = self._convert_to_grayscale(image)
        elif effect_type == ImageEffect.SEPIA:
            image = self._apply_sepia(image)
        elif effect_type == ImageEffect.INVERT:
            image.invertPixels()
        elif effect_type == ImageEffect.NONE:
            return True  # No effect

        # Save modified image temporarily
        temp_path = f"{image_path}_{effect_type.lower()}.png"
        image.save(temp_path)

        # Update document resource
        self.parent.document().addResource(QTextDocument.ResourceType.ImageResource,
                                          temp_path, image)

        # Update image format
        cursor = self.parent.textCursor()
        new_format = QTextImageFormat()
        new_format.setName(temp_path)
        new_format.setWidth(image.width())
        new_format.setHeight(image.height())

        cursor.setCharFormat(new_format)

        return True

    def adjust_brightness(self, factor):
        """Adjust image brightness (factor: -100 to 100)."""
        image_format = self.get_current_image()
        if not image_format:
            return False

        image_path = image_format.name()
        image = QImage(image_path)

        if image.isNull():
            return False

        # Adjust brightness
        adjusted_image = self._adjust_brightness(image, factor)

        # Save modified image temporarily
        temp_path = f"{image_path}_bright_{factor}.png"
        adjusted_image.save(temp_path)

        # Update document resource
        self.parent.document().addResource(QTextDocument.ResourceType.ImageResource,
                                          temp_path, adjusted_image)

        # Update image format
        cursor = self.parent.textCursor()
        new_format = QTextImageFormat()
        new_format.setName(temp_path)
        new_format.setWidth(adjusted_image.width())
        new_format.setHeight(adjusted_image.height())

        cursor.setCharFormat(new_format)

        return True

    def compress_image(self, quality=75):
        """Compress the current image (quality: 0-100)."""
        image_format = self.get_current_image()
        if not image_format:
            return False

        image_path = image_format.name()
        image = QImage(image_path)

        if image.isNull():
            return False

        # Compress and save
        compressed_path = f"{image_path}_compressed.jpg"
        image.save(compressed_path, "JPEG", quality)

        # Update document resource
        compressed_image = QImage(compressed_path)
        self.parent.document().addResource(QTextDocument.ResourceType.ImageResource,
                                          compressed_path, compressed_image)

        # Update image format
        cursor = self.parent.textCursor()
        new_format = QTextImageFormat()
        new_format.setName(compressed_path)
        new_format.setWidth(compressed_image.width())
        new_format.setHeight(compressed_image.height())

        cursor.setCharFormat(new_format)

        return True

    def reset_to_original(self):
        """Reset image to original size and appearance."""
        image_format = self.get_current_image()
        if not image_format:
            return False

        image_path = image_format.name()

        # Try to find original path (remove suffixes)
        original_path = image_path.split('_rotated')[0].split('_flipped')[0].split('_cropped')[0]

        if os.path.exists(original_path):
            image = QImage(original_path)
            if not image.isNull():
                # Update to original
                cursor = self.parent.textCursor()
                new_format = QTextImageFormat()
                new_format.setName(original_path)
                new_format.setWidth(image.width())
                new_format.setHeight(image.height())

                cursor.setCharFormat(new_format)
                return True

        return False

    # Helper methods for image processing
    def _convert_to_grayscale(self, image):
        """Convert image to grayscale."""
        for y in range(image.height()):
            for x in range(image.width()):
                pixel = image.pixel(x, y)
                gray = qGray(pixel)
                image.setPixel(x, y, qRgb(gray, gray, gray))
        return image

    def _apply_sepia(self, image):
        """Apply sepia tone to image."""
        for y in range(image.height()):
            for x in range(image.width()):
                color = QColor(image.pixel(x, y))
                r, g, b = color.red(), color.green(), color.blue()

                # Sepia formula
                tr = int(0.393 * r + 0.769 * g + 0.189 * b)
                tg = int(0.349 * r + 0.686 * g + 0.168 * b)
                tb = int(0.272 * r + 0.534 * g + 0.131 * b)

                # Clamp values
                tr = min(255, tr)
                tg = min(255, tg)
                tb = min(255, tb)

                image.setPixel(x, y, qRgb(tr, tg, tb))

        return image

    def _adjust_brightness(self, image, factor):
        """Adjust image brightness."""
        result = QImage(image)

        for y in range(result.height()):
            for x in range(result.width()):
                color = QColor(result.pixel(x, y))
                r = min(255, max(0, color.red() + factor))
                g = min(255, max(0, color.green() + factor))
                b = min(255, max(0, color.blue() + factor))

                result.setPixel(x, y, qRgb(r, g, b))

        return result


# Helper functions for Qt color operations
def qGray(rgb):
    """Convert RGB to grayscale value."""
    return (((rgb >> 16) & 0xff) * 11 + ((rgb >> 8) & 0xff) * 16 + (rgb & 0xff) * 5) // 32


def qRgb(r, g, b):
    """Create RGB value."""
    return (0xff << 24) | ((r & 0xff) << 16) | ((g & 0xff) << 8) | (b & 0xff)


class AdvancedImageDialog(QDialog):
    """Dialog for advanced image manipulation."""

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Advanced Image Editing")
        self.setModal(True)
        self.setMinimumWidth(500)

        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # Tab widget
        tabs = QTabWidget()

        # Size tab
        size_tab = self.create_size_tab()
        tabs.addTab(size_tab, "Size")

        # Rotate/Flip tab
        rotate_tab = self.create_rotate_tab()
        tabs.addTab(rotate_tab, "Rotate & Flip")

        # Effects tab
        effects_tab = self.create_effects_tab()
        tabs.addTab(effects_tab, "Effects")

        # Adjustments tab
        adjust_tab = self.create_adjust_tab()
        tabs.addTab(adjust_tab, "Adjustments")

        layout.addWidget(tabs)

        # Buttons
        button_layout = QHBoxLayout()

        reset_button = QPushButton("Reset to Original")
        reset_button.clicked.connect(self.reset_image)

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)

        button_layout.addWidget(reset_button)
        button_layout.addStretch()
        button_layout.addWidget(close_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def create_size_tab(self):
        """Create the size adjustment tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Size group
        size_group = QGroupBox("Image Size")
        size_layout = QFormLayout()

        self.width_spin = QSpinBox()
        self.width_spin.setRange(1, 10000)
        self.width_spin.setValue(400)
        self.width_spin.setSuffix(" px")

        self.height_spin = QSpinBox()
        self.height_spin.setRange(1, 10000)
        self.height_spin.setValue(300)
        self.height_spin.setSuffix(" px")

        self.maintain_aspect = QCheckBox("Maintain aspect ratio")
        self.maintain_aspect.setChecked(True)

        size_layout.addRow("Width:", self.width_spin)
        size_layout.addRow("Height:", self.height_spin)
        size_layout.addRow(self.maintain_aspect)

        apply_size_button = QPushButton("Apply Size")
        apply_size_button.clicked.connect(self.apply_size)
        size_layout.addRow(apply_size_button)

        size_group.setLayout(size_layout)
        layout.addWidget(size_group)

        # Scale group
        scale_group = QGroupBox("Scale")
        scale_layout = QFormLayout()

        self.scale_spin = QSpinBox()
        self.scale_spin.setRange(1, 500)
        self.scale_spin.setValue(100)
        self.scale_spin.setSuffix(" %")

        scale_layout.addRow("Scale:", self.scale_spin)

        apply_scale_button = QPushButton("Apply Scale")
        apply_scale_button.clicked.connect(self.apply_scale)
        scale_layout.addRow(apply_scale_button)

        scale_group.setLayout(scale_layout)
        layout.addWidget(scale_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_rotate_tab(self):
        """Create the rotate and flip tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Rotate group
        rotate_group = QGroupBox("Rotate")
        rotate_layout = QVBoxLayout()

        angle_layout = QHBoxLayout()
        angle_layout.addWidget(QLabel("Angle:"))
        self.angle_spin = QSpinBox()
        self.angle_spin.setRange(-360, 360)
        self.angle_spin.setValue(90)
        self.angle_spin.setSuffix("°")
        angle_layout.addWidget(self.angle_spin)
        rotate_layout.addLayout(angle_layout)

        button_layout = QHBoxLayout()
        rotate_button = QPushButton("Rotate")
        rotate_button.clicked.connect(self.rotate_image)
        rotate_90_button = QPushButton("Rotate 90° Right")
        rotate_90_button.clicked.connect(lambda: self.rotate_by_angle(90))
        rotate_270_button = QPushButton("Rotate 90° Left")
        rotate_270_button.clicked.connect(lambda: self.rotate_by_angle(-90))

        button_layout.addWidget(rotate_button)
        button_layout.addWidget(rotate_90_button)
        button_layout.addWidget(rotate_270_button)
        rotate_layout.addLayout(button_layout)

        rotate_group.setLayout(rotate_layout)
        layout.addWidget(rotate_group)

        # Flip group
        flip_group = QGroupBox("Flip")
        flip_layout = QHBoxLayout()

        flip_h_button = QPushButton("Flip Horizontal")
        flip_h_button.clicked.connect(lambda: self.flip_image(True, False))

        flip_v_button = QPushButton("Flip Vertical")
        flip_v_button.clicked.connect(lambda: self.flip_image(False, True))

        flip_layout.addWidget(flip_h_button)
        flip_layout.addWidget(flip_v_button)

        flip_group.setLayout(flip_layout)
        layout.addWidget(flip_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_effects_tab(self):
        """Create the effects tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        effects_group = QGroupBox("Image Effects")
        effects_layout = QVBoxLayout()

        self.effect_combo = QComboBox()
        self.effect_combo.addItems([
            ImageEffect.NONE,
            ImageEffect.GRAYSCALE,
            ImageEffect.SEPIA,
            ImageEffect.INVERT
        ])

        effects_layout.addWidget(QLabel("Choose an effect:"))
        effects_layout.addWidget(self.effect_combo)

        apply_effect_button = QPushButton("Apply Effect")
        apply_effect_button.clicked.connect(self.apply_effect)
        effects_layout.addWidget(apply_effect_button)

        effects_group.setLayout(effects_layout)
        layout.addWidget(effects_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_adjust_tab(self):
        """Create the adjustments tab."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Brightness
        brightness_group = QGroupBox("Brightness")
        brightness_layout = QVBoxLayout()

        self.brightness_slider = QSlider(Qt.Orientation.Horizontal)
        self.brightness_slider.setRange(-100, 100)
        self.brightness_slider.setValue(0)
        self.brightness_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.brightness_slider.setTickInterval(25)

        self.brightness_label = QLabel("0")
        self.brightness_slider.valueChanged.connect(
            lambda v: self.brightness_label.setText(str(v))
        )

        brightness_layout.addWidget(self.brightness_slider)
        brightness_layout.addWidget(self.brightness_label, alignment=Qt.AlignmentFlag.AlignCenter)

        apply_brightness_button = QPushButton("Apply Brightness")
        apply_brightness_button.clicked.connect(self.apply_brightness)
        brightness_layout.addWidget(apply_brightness_button)

        brightness_group.setLayout(brightness_layout)
        layout.addWidget(brightness_group)

        # Compression
        compression_group = QGroupBox("Image Compression")
        compression_layout = QFormLayout()

        self.quality_spin = QSpinBox()
        self.quality_spin.setRange(1, 100)
        self.quality_spin.setValue(75)
        self.quality_spin.setSuffix(" %")

        compression_layout.addRow("Quality:", self.quality_spin)

        compress_button = QPushButton("Compress Image")
        compress_button.clicked.connect(self.compress_image)
        compression_layout.addRow(compress_button)

        compression_group.setLayout(compression_layout)
        layout.addWidget(compression_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def apply_size(self):
        """Apply size changes."""
        width = self.width_spin.value()
        height = self.height_spin.value() if not self.maintain_aspect.isChecked() else None
        self.manager.resize_image(width, height, self.maintain_aspect.isChecked())

    def apply_scale(self):
        """Apply scale changes."""
        image_format = self.manager.get_current_image()
        if image_format:
            scale_factor = self.scale_spin.value() / 100.0
            current_width = image_format.width()
            new_width = int(current_width * scale_factor)
            self.manager.resize_image(new_width, None, True)

    def rotate_image(self):
        """Rotate image by specified angle."""
        angle = self.angle_spin.value()
        self.manager.rotate_image(angle)

    def rotate_by_angle(self, angle):
        """Rotate image by specific angle."""
        self.manager.rotate_image(angle)

    def flip_image(self, horizontal, vertical):
        """Flip the image."""
        self.manager.flip_image(horizontal, vertical)

    def apply_effect(self):
        """Apply the selected effect."""
        effect = self.effect_combo.currentText()
        self.manager.apply_effect(effect)

    def apply_brightness(self):
        """Apply brightness adjustment."""
        factor = self.brightness_slider.value()
        self.manager.adjust_brightness(factor)

    def compress_image(self):
        """Compress the image."""
        quality = self.quality_spin.value()
        self.manager.compress_image(quality)

    def reset_image(self):
        """Reset image to original."""
        if self.manager.reset_to_original():
            QMessageBox.information(self, "Success", "Image reset to original.")
        else:
            QMessageBox.warning(self, "Error", "Could not reset image to original.")
