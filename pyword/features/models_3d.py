"""3D Model support for PyWord - insert and manipulate 3D models."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                               QLabel, QGroupBox, QSlider, QSpinBox, QFileDialog,
                               QComboBox, QSplitter, QWidget, QCheckBox, QFormLayout)
from PySide6.QtCore import Qt, Signal, QRectF, QPointF, QSize, QTimer
from PySide6.QtGui import (QColor, QPainter, QPen, QBrush, QFont, QImage,
                          QTextImageFormat, QTextCursor, QMatrix4x4, QVector3D,
                          QQuaternion, QPainterPath, QLinearGradient)
from enum import Enum
from typing import Tuple, List
import math


class RenderMode(Enum):
    """3D model render modes."""
    WIREFRAME = "Wireframe"
    SOLID = "Solid"
    SHADED = "Shaded"
    TEXTURED = "Textured"


class Model3D:
    """Represents a 3D model."""

    def __init__(self, model_path: str = ""):
        self.model_path = model_path
        self.vertices = []
        self.faces = []
        self.position = QVector3D(0, 0, 0)
        self.rotation = QVector3D(0, 0, 0)  # Euler angles in degrees
        self.scale = QVector3D(1, 1, 1)
        self.color = QColor(100, 150, 200)
        self.render_mode = RenderMode.SHADED

        # Initialize with a simple cube if no model specified
        if not model_path:
            self.create_default_cube()

    def create_default_cube(self):
        """Create a simple cube model."""
        # Vertices of a unit cube
        self.vertices = [
            QVector3D(-1, -1, -1),  # 0
            QVector3D(1, -1, -1),   # 1
            QVector3D(1, 1, -1),    # 2
            QVector3D(-1, 1, -1),   # 3
            QVector3D(-1, -1, 1),   # 4
            QVector3D(1, -1, 1),    # 5
            QVector3D(1, 1, 1),     # 6
            QVector3D(-1, 1, 1),    # 7
        ]

        # Faces (triangles) - each face as two triangles
        self.faces = [
            # Front face
            [0, 1, 2], [0, 2, 3],
            # Back face
            [4, 6, 5], [4, 7, 6],
            # Top face
            [3, 2, 6], [3, 6, 7],
            # Bottom face
            [0, 5, 1], [0, 4, 5],
            # Right face
            [1, 5, 6], [1, 6, 2],
            # Left face
            [0, 3, 7], [0, 7, 4],
        ]

    def create_pyramid(self):
        """Create a pyramid model."""
        self.vertices = [
            QVector3D(0, 2, 0),     # Apex
            QVector3D(-1, 0, -1),   # Base
            QVector3D(1, 0, -1),
            QVector3D(1, 0, 1),
            QVector3D(-1, 0, 1),
        ]

        self.faces = [
            # Sides
            [0, 1, 2],
            [0, 2, 3],
            [0, 3, 4],
            [0, 4, 1],
            # Base
            [1, 3, 2], [1, 4, 3],
        ]

    def create_sphere(self, segments: int = 16):
        """Create a sphere model using UV sphere approach."""
        self.vertices = []
        self.faces = []

        # Create vertices
        for lat in range(segments + 1):
            theta = lat * math.pi / segments
            sin_theta = math.sin(theta)
            cos_theta = math.cos(theta)

            for lon in range(segments):
                phi = lon * 2 * math.pi / segments
                sin_phi = math.sin(phi)
                cos_phi = math.cos(phi)

                x = cos_phi * sin_theta
                y = cos_theta
                z = sin_phi * sin_theta

                self.vertices.append(QVector3D(x, y, z))

        # Create faces
        for lat in range(segments):
            for lon in range(segments):
                first = lat * segments + lon
                second = first + segments

                self.faces.append([first, second, first + 1])
                self.faces.append([second, second + 1, first + 1])

    def get_transform_matrix(self) -> QMatrix4x4:
        """Get the transformation matrix for this model."""
        matrix = QMatrix4x4()

        # Apply transformations in order: scale, rotate, translate
        matrix.translate(self.position)

        # Apply rotations (in order: Z, Y, X)
        matrix.rotate(self.rotation.z(), QVector3D(0, 0, 1))
        matrix.rotate(self.rotation.y(), QVector3D(0, 1, 0))
        matrix.rotate(self.rotation.x(), QVector3D(1, 0, 0))

        matrix.scale(self.scale)

        return matrix

    def load_from_file(self, file_path: str) -> bool:
        """Load a 3D model from a file.

        Note: In a real implementation, this would support formats like:
        - OBJ
        - STL
        - FBX
        - glTF

        For now, this is a placeholder.
        """
        self.model_path = file_path
        # Parse file and populate vertices and faces
        # This would use a 3D model loading library
        return True


class Model3DRenderer:
    """Renders 3D models to 2D images."""

    def __init__(self, width: int = 600, height: int = 600):
        self.width = width
        self.height = height
        self.camera_distance = 8.0
        self.fov = 60.0  # Field of view in degrees
        self.light_direction = QVector3D(1, 1, 1).normalized()

    def render_model(self, model: Model3D) -> QImage:
        """Render a 3D model to a 2D image."""
        image = QImage(self.width, self.height, QImage.Format_ARGB32)
        image.fill(Qt.white)

        painter = QPainter(image)
        painter.setRenderHint(QPainter.Antialiasing)

        # Set up view and projection
        view_matrix = self.get_view_matrix()
        proj_matrix = self.get_projection_matrix()
        model_matrix = model.get_transform_matrix()

        # Combined transformation matrix
        mvp_matrix = proj_matrix * view_matrix * model_matrix

        # Transform vertices
        transformed_vertices = []
        for vertex in model.vertices:
            # Transform to clip space
            transformed = mvp_matrix.map(vertex)

            # Perspective division
            if transformed.z() != 0:
                transformed.setX(transformed.x() / transformed.z())
                transformed.setY(transformed.y() / transformed.z())

            # Convert to screen space
            screen_x = (transformed.x() + 1) * self.width / 2
            screen_y = (1 - transformed.y()) * self.height / 2

            transformed_vertices.append((screen_x, screen_y, transformed.z()))

        # Render faces
        if model.render_mode == RenderMode.WIREFRAME:
            self.render_wireframe(painter, model, transformed_vertices)
        elif model.render_mode == RenderMode.SOLID:
            self.render_solid(painter, model, transformed_vertices)
        elif model.render_mode == RenderMode.SHADED:
            self.render_shaded(painter, model, transformed_vertices, model_matrix)
        else:
            self.render_shaded(painter, model, transformed_vertices, model_matrix)

        painter.end()
        return image

    def get_view_matrix(self) -> QMatrix4x4:
        """Get the view matrix (camera transformation)."""
        matrix = QMatrix4x4()
        matrix.lookAt(
            QVector3D(0, 0, self.camera_distance),  # Camera position
            QVector3D(0, 0, 0),  # Look at origin
            QVector3D(0, 1, 0)   # Up vector
        )
        return matrix

    def get_projection_matrix(self) -> QMatrix4x4:
        """Get the projection matrix (perspective)."""
        matrix = QMatrix4x4()
        aspect = self.width / self.height
        matrix.perspective(self.fov, aspect, 0.1, 100.0)
        return matrix

    def render_wireframe(self, painter: QPainter, model: Model3D,
                        vertices: List[Tuple[float, float, float]]):
        """Render model as wireframe."""
        painter.setPen(QPen(model.color, 2))

        for face in model.faces:
            # Draw edges of the triangle
            for i in range(3):
                v1 = vertices[face[i]]
                v2 = vertices[face[(i + 1) % 3]]

                painter.drawLine(QPointF(v1[0], v1[1]),
                               QPointF(v2[0], v2[1]))

    def render_solid(self, painter: QPainter, model: Model3D,
                    vertices: List[Tuple[float, float, float]]):
        """Render model as solid color."""
        painter.setPen(Qt.NoPen)
        painter.setBrush(model.color)

        # Sort faces by depth (painter's algorithm)
        sorted_faces = self.sort_faces_by_depth(model.faces, vertices)

        for face in sorted_faces:
            points = [QPointF(vertices[i][0], vertices[i][1]) for i in face]
            painter.drawPolygon(points)

    def render_shaded(self, painter: QPainter, model: Model3D,
                     vertices: List[Tuple[float, float, float]],
                     model_matrix: QMatrix4x4):
        """Render model with basic shading."""
        # Sort faces by depth
        sorted_faces = self.sort_faces_by_depth(model.faces, vertices)

        for face in sorted_faces:
            # Calculate face normal in world space
            v0 = model.vertices[face[0]]
            v1 = model.vertices[face[1]]
            v2 = model.vertices[face[2]]

            # Transform vertices to world space
            v0_world = model_matrix.map(v0)
            v1_world = model_matrix.map(v1)
            v2_world = model_matrix.map(v2)

            # Calculate normal
            edge1 = v1_world - v0_world
            edge2 = v2_world - v0_world
            normal = QVector3D.crossProduct(edge1, edge2).normalized()

            # Calculate lighting (simple diffuse)
            light_intensity = max(0, QVector3D.dotProduct(normal, self.light_direction))
            light_intensity = 0.3 + 0.7 * light_intensity  # Ambient + diffuse

            # Adjust color based on lighting
            r = int(model.color.red() * light_intensity)
            g = int(model.color.green() * light_intensity)
            b = int(model.color.blue() * light_intensity)
            face_color = QColor(r, g, b)

            # Draw face
            points = [QPointF(vertices[i][0], vertices[i][1]) for i in face]
            painter.setPen(QPen(face_color.darker(120), 1))
            painter.setBrush(face_color)
            painter.drawPolygon(points)

    def sort_faces_by_depth(self, faces: List[List[int]],
                           vertices: List[Tuple[float, float, float]]) -> List[List[int]]:
        """Sort faces by average depth (for painter's algorithm)."""
        def face_depth(face):
            return sum(vertices[i][2] for i in face) / len(face)

        return sorted(faces, key=face_depth, reverse=True)


class Model3DEditor(QDialog):
    """Dialog for inserting and configuring 3D models."""

    model_inserted = Signal(Model3D)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert 3D Model")
        self.resize(1000, 700)

        self.model = Model3D()
        self.renderer = Model3DRenderer()

        # Auto-rotation for preview
        self.auto_rotate = False
        self.rotation_timer = QTimer(self)
        self.rotation_timer.timeout.connect(self.auto_rotate_step)

        self.setup_ui()
        self.update_preview()

    def setup_ui(self):
        """Initialize the 3D model editor UI."""
        layout = QVBoxLayout(self)

        # Splitter for controls and preview
        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter)

        # Left side: Controls
        controls_widget = QWidget()
        controls_layout = QVBoxLayout(controls_widget)

        # Model selection
        model_group = QGroupBox("Model")
        model_layout = QVBoxLayout(model_group)

        load_button = QPushButton("Load from File...")
        load_button.clicked.connect(self.load_model)
        model_layout.addWidget(load_button)

        model_layout.addWidget(QLabel("or choose a preset:"))

        self.preset_combo = QComboBox()
        self.preset_combo.addItems(["Cube", "Pyramid", "Sphere"])
        self.preset_combo.currentIndexChanged.connect(self.on_preset_changed)
        model_layout.addWidget(self.preset_combo)

        controls_layout.addWidget(model_group)

        # Rotation controls
        rotation_group = QGroupBox("Rotation")
        rotation_layout = QFormLayout(rotation_group)

        self.rotation_x_slider = QSlider(Qt.Horizontal)
        self.rotation_x_slider.setRange(-180, 180)
        self.rotation_x_slider.setValue(30)
        self.rotation_x_slider.valueChanged.connect(self.on_rotation_changed)
        rotation_layout.addRow("X Axis:", self.rotation_x_slider)

        self.rotation_y_slider = QSlider(Qt.Horizontal)
        self.rotation_y_slider.setRange(-180, 180)
        self.rotation_y_slider.setValue(45)
        self.rotation_y_slider.valueChanged.connect(self.on_rotation_changed)
        rotation_layout.addRow("Y Axis:", self.rotation_y_slider)

        self.rotation_z_slider = QSlider(Qt.Horizontal)
        self.rotation_z_slider.setRange(-180, 180)
        self.rotation_z_slider.setValue(0)
        self.rotation_z_slider.valueChanged.connect(self.on_rotation_changed)
        rotation_layout.addRow("Z Axis:", self.rotation_z_slider)

        self.auto_rotate_check = QCheckBox("Auto-rotate")
        self.auto_rotate_check.stateChanged.connect(self.on_auto_rotate_changed)
        rotation_layout.addRow("", self.auto_rotate_check)

        controls_layout.addWidget(rotation_group)

        # Scale controls
        scale_group = QGroupBox("Scale")
        scale_layout = QFormLayout(scale_group)

        self.scale_slider = QSlider(Qt.Horizontal)
        self.scale_slider.setRange(10, 200)
        self.scale_slider.setValue(100)
        self.scale_slider.valueChanged.connect(self.on_scale_changed)
        scale_layout.addRow("Size:", self.scale_slider)

        controls_layout.addWidget(scale_group)

        # Render mode
        render_group = QGroupBox("Render Mode")
        render_layout = QVBoxLayout(render_group)

        self.render_combo = QComboBox()
        for mode in RenderMode:
            self.render_combo.addItem(mode.value, mode)
        self.render_combo.setCurrentIndex(2)  # Default to Shaded
        self.render_combo.currentIndexChanged.connect(self.on_render_mode_changed)
        render_layout.addWidget(self.render_combo)

        controls_layout.addWidget(render_group)

        # Color
        color_group = QGroupBox("Color")
        color_layout = QHBoxLayout(color_group)

        self.color_button = QPushButton("Choose Color...")
        self.color_button.clicked.connect(self.choose_color)
        color_layout.addWidget(self.color_button)

        controls_layout.addWidget(color_group)

        controls_layout.addStretch()

        splitter.addWidget(controls_widget)

        # Right side: Preview
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)

        preview_label = QLabel("Preview:")
        preview_layout.addWidget(preview_label)

        self.preview_label = QLabel()
        self.preview_label.setMinimumSize(600, 600)
        self.preview_label.setStyleSheet("border: 1px solid #d0d0d0; background: white;")
        self.preview_label.setAlignment(Qt.AlignCenter)
        preview_layout.addWidget(self.preview_label)

        splitter.addWidget(preview_widget)

        # Dialog buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        insert_button = QPushButton("Insert")
        insert_button.clicked.connect(self.accept)
        button_layout.addWidget(insert_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        # Initialize rotation
        self.on_rotation_changed()

    def load_model(self):
        """Load a 3D model from file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Load 3D Model",
            "",
            "3D Models (*.obj *.stl *.fbx *.gltf);;All Files (*.*)"
        )

        if file_path:
            self.model.load_from_file(file_path)
            self.update_preview()

    def on_preset_changed(self, index):
        """Handle preset model selection."""
        if index == 0:  # Cube
            self.model.create_default_cube()
        elif index == 1:  # Pyramid
            self.model.create_pyramid()
        elif index == 2:  # Sphere
            self.model.create_sphere()

        self.update_preview()

    def on_rotation_changed(self):
        """Handle rotation slider changes."""
        self.model.rotation = QVector3D(
            self.rotation_x_slider.value(),
            self.rotation_y_slider.value(),
            self.rotation_z_slider.value()
        )
        self.update_preview()

    def on_scale_changed(self):
        """Handle scale slider changes."""
        scale_factor = self.scale_slider.value() / 100.0
        self.model.scale = QVector3D(scale_factor, scale_factor, scale_factor)
        self.update_preview()

    def on_render_mode_changed(self):
        """Handle render mode change."""
        self.model.render_mode = self.render_combo.currentData()
        self.update_preview()

    def on_auto_rotate_changed(self, state):
        """Handle auto-rotate checkbox."""
        self.auto_rotate = (state == Qt.Checked)
        if self.auto_rotate:
            self.rotation_timer.start(50)  # Update every 50ms
        else:
            self.rotation_timer.stop()

    def auto_rotate_step(self):
        """Perform one step of auto-rotation."""
        current_y = self.rotation_y_slider.value()
        new_y = (current_y + 2) % 360
        if new_y > 180:
            new_y -= 360
        self.rotation_y_slider.setValue(new_y)

    def choose_color(self):
        """Choose model color."""
        from PySide6.QtWidgets import QColorDialog
        color = QColorDialog.getColor(self.model.color, self)
        if color.isValid():
            self.model.color = color
            self.update_preview()

    def update_preview(self):
        """Update the 3D model preview."""
        image = self.renderer.render_model(self.model)
        self.preview_label.setPixmap(image.scaled(
            self.preview_label.size(),
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        ))

    def get_model(self) -> Model3D:
        """Get the configured 3D model."""
        return self.model


class Model3DManager:
    """Manages 3D models in a document."""

    def __init__(self, editor):
        self.editor = editor
        self.models = {}
        self.renderer = Model3DRenderer()

    def show_model_editor(self):
        """Show the 3D model editor dialog."""
        dialog = Model3DEditor(self.editor)
        if dialog.exec() == QDialog.Accepted:
            model = dialog.get_model()
            self.insert_model(model)

    def insert_model(self, model: Model3D):
        """Insert a 3D model into the document."""
        # Generate unique ID
        model_id = f"model3d_{len(self.models)}"
        self.models[model_id] = model

        # Render model
        image = self.renderer.render_model(model)

        # Insert into document
        cursor = self.editor.textCursor()

        image_format = QTextImageFormat()
        image_format.setName(model_id)

        self.editor.document().addResource(
            QTextCursor.ImageResource,
            model_id,
            image
        )

        cursor.insertImage(image_format)
