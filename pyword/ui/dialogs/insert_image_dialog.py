"""Insert Image Dialog for PyWord."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QLineEdit, QFileDialog, QGroupBox,
                             QFormLayout, QSpinBox, QComboBox)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from .base_dialog import BaseDialog


class InsertImageDialog(BaseDialog):
    """Dialog for inserting an image."""

    def __init__(self, parent=None):
        super().__init__("Insert Image", parent)
        self.image_path = ""
        self.setup_ui()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout()

        # File selection
        file_group = QGroupBox("Image File")
        file_layout = QHBoxLayout()

        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        file_layout.addWidget(self.file_edit)

        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.browse_button)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Image preview
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()

        self.preview_label = QLabel()
        self.preview_label.setMinimumSize(300, 200)
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("border: 1px solid gray;")
        self.preview_label.setText("No image selected")
        preview_layout.addWidget(self.preview_label)

        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)

        # Size options
        size_group = QGroupBox("Size")
        size_layout = QFormLayout()

        self.width_spinbox = QSpinBox()
        self.width_spinbox.setMinimum(10)
        self.width_spinbox.setMaximum(2000)
        self.width_spinbox.setValue(300)
        self.width_spinbox.setSuffix(" px")
        size_layout.addRow("Width:", self.width_spinbox)

        self.height_spinbox = QSpinBox()
        self.height_spinbox.setMinimum(10)
        self.height_spinbox.setMaximum(2000)
        self.height_spinbox.setValue(200)
        self.height_spinbox.setSuffix(" px")
        size_layout.addRow("Height:", self.height_spinbox)

        size_group.setLayout(size_layout)
        layout.addWidget(size_group)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.ok_button = QPushButton("Insert")
        self.ok_button.clicked.connect(self.accept)
        self.ok_button.setEnabled(False)
        button_layout.addWidget(self.ok_button)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.resize(400, 500)

    def browse_file(self):
        """Browse for an image file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Image",
            "",
            "Image Files (*.png *.jpg *.jpeg *.gif *.bmp *.svg);;All Files (*.*)"
        )

        if file_path:
            self.image_path = file_path
            self.file_edit.setText(file_path)
            self.ok_button.setEnabled(True)

            # Show preview
            pixmap = QPixmap(file_path)
            if not pixmap.isNull():
                scaled_pixmap = pixmap.scaled(
                    self.preview_label.size(),
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                )
                self.preview_label.setPixmap(scaled_pixmap)

                # Update size spinboxes
                self.width_spinbox.setValue(pixmap.width())
                self.height_spinbox.setValue(pixmap.height())

    def get_image_path(self):
        """Get the selected image path."""
        return self.image_path

    def get_width(self):
        """Get the image width."""
        return self.width_spinbox.value()

    def get_height(self):
        """Get the image height."""
        return self.height_spinbox.value()
