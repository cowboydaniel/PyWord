"""SmartArt and Diagram tools for PyWord - create professional diagrams."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                               QListWidget, QListWidgetItem, QLabel, QGroupBox,
                               QTextEdit, QLineEdit, QSplitter, QWidget, QScrollArea,
                               QGridLayout, QColorDialog, QSpinBox, QComboBox)
from PySide6.QtCore import Qt, Signal, QRectF, QPointF, QSize
from PySide6.QtGui import (QColor, QPainter, QPen, QBrush, QFont, QImage,
                          QTextImageFormat, QTextCursor, QPainterPath, QLinearGradient,
                          QPolygonF)
from enum import Enum
from typing import List, Dict, Tuple


class SmartArtType(Enum):
    """Types of SmartArt diagrams."""
    LIST = "List"
    PROCESS = "Process"
    CYCLE = "Cycle"
    HIERARCHY = "Hierarchy"
    RELATIONSHIP = "Relationship"
    MATRIX = "Matrix"
    PYRAMID = "Pyramid"


class SmartArtStyle(Enum):
    """Visual styles for SmartArt."""
    SIMPLE = "Simple"
    POLISHED = "Polished"
    INTENSE = "Intense"
    INSET = "Inset"


class DiagramNode:
    """Represents a node in a diagram."""

    def __init__(self, text: str = "", level: int = 0):
        self.text = text
        self.level = level
        self.children = []
        self.color = QColor(68, 114, 196)
        self.position = QPointF(0, 0)
        self.size = QSize(120, 60)

    def add_child(self, node: 'DiagramNode'):
        """Add a child node."""
        self.children.append(node)

    def remove_child(self, node: 'DiagramNode'):
        """Remove a child node."""
        if node in self.children:
            self.children.remove(node)


class SmartArtDiagram:
    """Represents a SmartArt diagram."""

    def __init__(self, diagram_type: SmartArtType = SmartArtType.PROCESS):
        self.diagram_type = diagram_type
        self.style = SmartArtStyle.POLISHED
        self.nodes = []
        self.title = ""
        self.color_scheme = self.default_color_scheme()

    def default_color_scheme(self) -> List[QColor]:
        """Get default color scheme."""
        return [
            QColor(68, 114, 196),   # Blue
            QColor(237, 125, 49),   # Orange
            QColor(165, 165, 165),  # Gray
            QColor(255, 192, 0),    # Yellow
            QColor(91, 155, 213),   # Light Blue
            QColor(112, 173, 71),   # Green
        ]

    def add_node(self, text: str, level: int = 0) -> DiagramNode:
        """Add a node to the diagram."""
        node = DiagramNode(text, level)
        node.color = self.color_scheme[len(self.nodes) % len(self.color_scheme)]
        self.nodes.append(node)
        return node

    def remove_node(self, node: DiagramNode):
        """Remove a node from the diagram."""
        if node in self.nodes:
            self.nodes.remove(node)


class SmartArtRenderer:
    """Renders SmartArt diagrams to images."""

    def __init__(self, width: int = 800, height: int = 600):
        self.width = width
        self.height = height
        self.margin = 50

    def render_diagram(self, diagram: SmartArtDiagram) -> QImage:
        """Render a SmartArt diagram to an image."""
        image = QImage(self.width, self.height, QImage.Format_ARGB32)
        image.fill(Qt.white)

        painter = QPainter(image)
        painter.setRenderHint(QPainter.Antialiasing)

        # Render based on diagram type
        if diagram.diagram_type == SmartArtType.LIST:
            self.render_list_diagram(painter, diagram)
        elif diagram.diagram_type == SmartArtType.PROCESS:
            self.render_process_diagram(painter, diagram)
        elif diagram.diagram_type == SmartArtType.CYCLE:
            self.render_cycle_diagram(painter, diagram)
        elif diagram.diagram_type == SmartArtType.HIERARCHY:
            self.render_hierarchy_diagram(painter, diagram)
        elif diagram.diagram_type == SmartArtType.RELATIONSHIP:
            self.render_relationship_diagram(painter, diagram)
        elif diagram.diagram_type == SmartArtType.MATRIX:
            self.render_matrix_diagram(painter, diagram)
        elif diagram.diagram_type == SmartArtType.PYRAMID:
            self.render_pyramid_diagram(painter, diagram)

        painter.end()
        return image

    def render_list_diagram(self, painter: QPainter, diagram: SmartArtDiagram):
        """Render a list-style diagram."""
        if not diagram.nodes:
            return

        # Calculate layout
        node_height = 80
        node_width = self.width - 2 * self.margin
        spacing = 20

        y = self.margin

        for i, node in enumerate(diagram.nodes):
            # Draw node background
            rect = QRectF(self.margin, y, node_width, node_height)

            # Create gradient
            gradient = QLinearGradient(rect.topLeft(), rect.bottomLeft())
            gradient.setColorAt(0, node.color)
            gradient.setColorAt(1, node.color.darker(120))

            painter.fillRect(rect, gradient)
            painter.setPen(QPen(node.color.darker(150), 2))
            painter.drawRect(rect)

            # Draw text
            painter.setPen(Qt.white)
            painter.setFont(QFont("Arial", 12, QFont.Weight.Bold))
            painter.drawText(rect, Qt.AlignCenter, node.text)

            y += node_height + spacing

    def render_process_diagram(self, painter: QPainter, diagram: SmartArtDiagram):
        """Render a process flow diagram."""
        if not diagram.nodes:
            return

        num_nodes = len(diagram.nodes)
        node_width = 150
        node_height = 80
        spacing = 40

        # Calculate starting position to center the diagram
        total_width = num_nodes * node_width + (num_nodes - 1) * spacing
        x_start = (self.width - total_width) / 2
        y = (self.height - node_height) / 2

        for i, node in enumerate(diagram.nodes):
            x = x_start + i * (node_width + spacing)
            rect = QRectF(x, y, node_width, node_height)

            # Draw shape with rounded corners
            gradient = QLinearGradient(rect.topLeft(), rect.bottomLeft())
            gradient.setColorAt(0, node.color.lighter(110))
            gradient.setColorAt(1, node.color)

            painter.setBrush(gradient)
            painter.setPen(QPen(node.color.darker(150), 2))
            painter.drawRoundedRect(rect, 10, 10)

            # Draw text
            painter.setPen(Qt.white)
            painter.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            painter.drawText(rect, Qt.AlignCenter | Qt.TextWordWrap, node.text)

            # Draw arrow to next node
            if i < num_nodes - 1:
                arrow_start = QPointF(x + node_width, y + node_height / 2)
                arrow_end = QPointF(x + node_width + spacing, y + node_height / 2)

                painter.setPen(QPen(QColor(100, 100, 100), 3))
                painter.drawLine(arrow_start, arrow_end)

                # Draw arrow head
                self.draw_arrow_head(painter, arrow_end, 0)

    def render_cycle_diagram(self, painter: QPainter, diagram: SmartArtDiagram):
        """Render a cycle diagram."""
        if not diagram.nodes:
            return

        num_nodes = len(diagram.nodes)
        center_x = self.width / 2
        center_y = self.height / 2
        radius = min(self.width, self.height) / 2 - 100

        node_width = 120
        node_height = 80

        for i, node in enumerate(diagram.nodes):
            # Calculate position on circle
            angle = 2 * 3.14159 * i / num_nodes - 3.14159 / 2  # Start from top
            x = center_x + radius * (angle ** 0 if i == 0 else 1) * (1 if i == 0 else
                                                                     (2.71828 ** (1j * angle)).real)
            y = center_y + radius * (0 if i == 0 else (2.71828 ** (1j * angle)).imag)

            # Simplified calculation
            import math
            x = center_x + radius * math.cos(angle)
            y = center_y + radius * math.sin(angle)

            rect = QRectF(x - node_width / 2, y - node_height / 2,
                         node_width, node_height)

            # Draw node
            painter.setBrush(node.color)
            painter.setPen(QPen(node.color.darker(150), 2))
            painter.drawEllipse(rect)

            # Draw text
            painter.setPen(Qt.white)
            painter.setFont(QFont("Arial", 9, QFont.Weight.Bold))
            painter.drawText(rect, Qt.AlignCenter | Qt.TextWordWrap, node.text)

            # Draw arrow to next node
            if num_nodes > 1:
                next_i = (i + 1) % num_nodes
                next_angle = 2 * 3.14159 * next_i / num_nodes - 3.14159 / 2
                next_x = center_x + radius * math.cos(next_angle)
                next_y = center_y + radius * math.sin(next_angle)

                # Draw curved arrow
                painter.setPen(QPen(QColor(100, 100, 100), 2))
                painter.drawLine(QPointF(x, y), QPointF(next_x, next_y))

    def render_hierarchy_diagram(self, painter: QPainter, diagram: SmartArtDiagram):
        """Render an organizational hierarchy diagram."""
        if not diagram.nodes:
            return

        node_width = 120
        node_height = 60
        h_spacing = 40
        v_spacing = 60

        # Simple tree layout (top node, then children below)
        # For a full implementation, you'd calculate proper tree positions

        # Draw root at top
        if diagram.nodes:
            root = diagram.nodes[0]
            root_x = (self.width - node_width) / 2
            root_y = self.margin

            rect = QRectF(root_x, root_y, node_width, node_height)
            self.draw_hierarchy_node(painter, rect, root)

            # Draw child nodes
            num_children = len(diagram.nodes) - 1
            if num_children > 0:
                total_width = num_children * node_width + (num_children - 1) * h_spacing
                child_x_start = (self.width - total_width) / 2
                child_y = root_y + node_height + v_spacing

                for i, child in enumerate(diagram.nodes[1:]):
                    child_x = child_x_start + i * (node_width + h_spacing)
                    child_rect = QRectF(child_x, child_y, node_width, node_height)
                    self.draw_hierarchy_node(painter, child_rect, child)

                    # Draw connecting line
                    painter.setPen(QPen(QColor(100, 100, 100), 2))
                    painter.drawLine(
                        QPointF(root_x + node_width / 2, root_y + node_height),
                        QPointF(child_x + node_width / 2, child_y)
                    )

    def draw_hierarchy_node(self, painter: QPainter, rect: QRectF, node: DiagramNode):
        """Draw a single hierarchy node."""
        gradient = QLinearGradient(rect.topLeft(), rect.bottomLeft())
        gradient.setColorAt(0, node.color.lighter(120))
        gradient.setColorAt(1, node.color)

        painter.setBrush(gradient)
        painter.setPen(QPen(node.color.darker(150), 2))
        painter.drawRoundedRect(rect, 5, 5)

        painter.setPen(Qt.white)
        painter.setFont(QFont("Arial", 9, QFont.Weight.Bold))
        painter.drawText(rect, Qt.AlignCenter | Qt.TextWordWrap, node.text)

    def render_relationship_diagram(self, painter: QPainter, diagram: SmartArtDiagram):
        """Render a relationship diagram (Venn diagram style)."""
        if not diagram.nodes:
            return

        center_x = self.width / 2
        center_y = self.height / 2
        circle_radius = 120

        if len(diagram.nodes) == 1:
            # Single circle
            node = diagram.nodes[0]
            painter.setBrush(QColor(node.color.red(), node.color.green(),
                                   node.color.blue(), 128))
            painter.setPen(QPen(node.color.darker(150), 2))
            painter.drawEllipse(QPointF(center_x, center_y),
                              circle_radius, circle_radius)

            painter.setPen(Qt.black)
            painter.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            painter.drawText(QRectF(center_x - 60, center_y - 10, 120, 20),
                           Qt.AlignCenter, node.text)

        elif len(diagram.nodes) >= 2:
            # Two overlapping circles
            offset = circle_radius * 0.7

            for i, node in enumerate(diagram.nodes[:3]):  # Max 3 circles
                if i == 0:
                    x, y = center_x - offset, center_y
                elif i == 1:
                    x, y = center_x + offset, center_y
                else:
                    x, y = center_x, center_y + offset

                color = QColor(node.color.red(), node.color.green(),
                             node.color.blue(), 100)
                painter.setBrush(color)
                painter.setPen(QPen(node.color.darker(150), 2))
                painter.drawEllipse(QPointF(x, y), circle_radius, circle_radius)

                # Draw label
                painter.setPen(Qt.black)
                painter.setFont(QFont("Arial", 9, QFont.Weight.Bold))
                label_rect = QRectF(x - 50, y - circle_radius - 20, 100, 20)
                painter.drawText(label_rect, Qt.AlignCenter, node.text)

    def render_matrix_diagram(self, painter: QPainter, diagram: SmartArtDiagram):
        """Render a matrix diagram (2x2 grid)."""
        if not diagram.nodes:
            return

        margin = 100
        grid_width = self.width - 2 * margin
        grid_height = self.height - 2 * margin

        cell_width = grid_width / 2
        cell_height = grid_height / 2

        # Draw up to 4 quadrants
        positions = [
            (margin, margin),  # Top-left
            (margin + cell_width, margin),  # Top-right
            (margin, margin + cell_height),  # Bottom-left
            (margin + cell_width, margin + cell_height)  # Bottom-right
        ]

        for i, (x, y) in enumerate(positions):
            if i >= len(diagram.nodes):
                break

            node = diagram.nodes[i]
            rect = QRectF(x + 10, y + 10, cell_width - 20, cell_height - 20)

            # Draw cell
            gradient = QLinearGradient(rect.topLeft(), rect.bottomRight())
            gradient.setColorAt(0, node.color.lighter(120))
            gradient.setColorAt(1, node.color)

            painter.setBrush(gradient)
            painter.setPen(QPen(node.color.darker(150), 2))
            painter.drawRect(rect)

            # Draw text
            painter.setPen(Qt.white)
            painter.setFont(QFont("Arial", 11, QFont.Weight.Bold))
            painter.drawText(rect, Qt.AlignCenter | Qt.TextWordWrap, node.text)

    def render_pyramid_diagram(self, painter: QPainter, diagram: SmartArtDiagram):
        """Render a pyramid diagram."""
        if not diagram.nodes:
            return

        margin = 100
        pyramid_width = self.width - 2 * margin
        pyramid_height = self.height - 2 * margin

        num_levels = len(diagram.nodes)
        level_height = pyramid_height / num_levels

        for i, node in enumerate(diagram.nodes):
            # Calculate trapezoid for this level
            level_y = margin + i * level_height
            top_width = pyramid_width * (num_levels - i) / num_levels
            bottom_width = pyramid_width * (num_levels - i - 1) / num_levels if i < num_levels - 1 else pyramid_width

            top_left = QPointF((self.width - top_width) / 2, level_y)
            top_right = QPointF((self.width + top_width) / 2, level_y)
            bottom_right = QPointF((self.width + bottom_width) / 2,
                                  level_y + level_height)
            bottom_left = QPointF((self.width - bottom_width) / 2,
                                 level_y + level_height)

            # Draw trapezoid
            polygon = QPolygonF([top_left, top_right, bottom_right, bottom_left])

            gradient = QLinearGradient(top_left, bottom_left)
            gradient.setColorAt(0, node.color.lighter(110))
            gradient.setColorAt(1, node.color)

            painter.setBrush(gradient)
            painter.setPen(QPen(node.color.darker(150), 2))
            painter.drawPolygon(polygon)

            # Draw text
            text_rect = QRectF((self.width - top_width) / 2 + 10, level_y + 5,
                             top_width - 20, level_height - 10)
            painter.setPen(Qt.white)
            painter.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            painter.drawText(text_rect, Qt.AlignCenter | Qt.TextWordWrap, node.text)

    def draw_arrow_head(self, painter: QPainter, tip: QPointF, angle: float):
        """Draw an arrow head."""
        import math
        arrow_size = 10

        # Calculate arrow points
        p1 = QPointF(
            tip.x() - arrow_size * math.cos(angle - math.pi / 6),
            tip.y() - arrow_size * math.sin(angle - math.pi / 6)
        )
        p2 = QPointF(
            tip.x() - arrow_size * math.cos(angle + math.pi / 6),
            tip.y() - arrow_size * math.sin(angle + math.pi / 6)
        )

        polygon = QPolygonF([tip, p1, p2])
        painter.setBrush(painter.pen().color())
        painter.drawPolygon(polygon)


class SmartArtEditor(QDialog):
    """Dialog for creating and editing SmartArt diagrams."""

    diagram_created = Signal(SmartArtDiagram)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert SmartArt")
        self.resize(1000, 700)

        self.diagram = SmartArtDiagram()
        self.renderer = SmartArtRenderer()

        # Add default nodes
        self.diagram.add_node("First")
        self.diagram.add_node("Second")
        self.diagram.add_node("Third")

        self.setup_ui()
        self.update_preview()

    def setup_ui(self):
        """Initialize the SmartArt editor UI."""
        layout = QVBoxLayout(self)

        # Splitter for settings and preview
        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter)

        # Left side: Settings
        settings_widget = QWidget()
        settings_layout = QVBoxLayout(settings_widget)

        # Diagram type selection
        type_group = QGroupBox("Diagram Type")
        type_layout = QVBoxLayout(type_group)

        self.type_list = QListWidget()
        for diagram_type in SmartArtType:
            item = QListWidgetItem(diagram_type.value)
            item.setData(Qt.UserRole, diagram_type)
            self.type_list.addItem(item)
        self.type_list.setCurrentRow(1)  # Default to Process
        self.type_list.currentItemChanged.connect(self.on_type_changed)
        type_layout.addWidget(self.type_list)

        settings_layout.addWidget(type_group)

        # Text input
        text_group = QGroupBox("Diagram Content")
        text_layout = QVBoxLayout(text_group)

        text_layout.addWidget(QLabel("Enter text for each item (one per line):"))

        self.text_input = QTextEdit()
        self.text_input.setPlainText("First\nSecond\nThird")
        self.text_input.textChanged.connect(self.on_text_changed)
        text_layout.addWidget(self.text_input)

        settings_layout.addWidget(text_group)

        # Color scheme
        color_group = QGroupBox("Colors")
        color_layout = QHBoxLayout(color_group)

        self.color_combo = QComboBox()
        self.color_combo.addItems(["Blue", "Orange", "Green", "Red", "Purple"])
        self.color_combo.currentIndexChanged.connect(self.on_color_changed)
        color_layout.addWidget(self.color_combo)

        settings_layout.addWidget(color_group)

        splitter.addWidget(settings_widget)

        # Right side: Preview
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)

        preview_label = QLabel("Preview:")
        preview_layout.addWidget(preview_label)

        self.preview_label = QLabel()
        self.preview_label.setMinimumSize(700, 500)
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

    def on_type_changed(self, current, previous):
        """Handle diagram type change."""
        if current:
            diagram_type = current.data(Qt.UserRole)
            self.diagram.diagram_type = diagram_type
            self.update_preview()

    def on_text_changed(self):
        """Handle text input change."""
        text = self.text_input.toPlainText()
        lines = [line.strip() for line in text.split('\n') if line.strip()]

        # Rebuild nodes
        self.diagram.nodes.clear()
        for line in lines:
            self.diagram.add_node(line)

        self.update_preview()

    def on_color_changed(self, index):
        """Handle color scheme change."""
        color_schemes = {
            0: [QColor(68, 114, 196), QColor(91, 155, 213), QColor(142, 180, 227)],  # Blue
            1: [QColor(237, 125, 49), QColor(244, 176, 132), QColor(250, 210, 186)],  # Orange
            2: [QColor(112, 173, 71), QColor(146, 205, 103), QColor(196, 226, 166)],  # Green
            3: [QColor(192, 0, 0), QColor(230, 92, 92), QColor(244, 166, 166)],  # Red
            4: [QColor(112, 48, 160), QColor(155, 107, 196), QColor(194, 165, 216)],  # Purple
        }

        if index in color_schemes:
            self.diagram.color_scheme = color_schemes[index]
            # Reapply colors to nodes
            for i, node in enumerate(self.diagram.nodes):
                node.color = self.diagram.color_scheme[i % len(self.diagram.color_scheme)]
            self.update_preview()

    def update_preview(self):
        """Update the diagram preview."""
        image = self.renderer.render_diagram(self.diagram)
        self.preview_label.setPixmap(image.scaled(
            self.preview_label.size(),
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        ))

    def get_diagram(self) -> SmartArtDiagram:
        """Get the current diagram."""
        return self.diagram


class SmartArtManager:
    """Manages SmartArt diagrams in a document."""

    def __init__(self, editor):
        self.editor = editor
        self.diagrams = {}
        self.renderer = SmartArtRenderer()

    def show_smartart_editor(self):
        """Show the SmartArt editor dialog."""
        dialog = SmartArtEditor(self.editor)
        if dialog.exec() == QDialog.Accepted:
            diagram = dialog.get_diagram()
            self.insert_diagram(diagram)

    def insert_diagram(self, diagram: SmartArtDiagram):
        """Insert a SmartArt diagram into the document."""
        # Generate unique ID
        diagram_id = f"smartart_{len(self.diagrams)}"
        self.diagrams[diagram_id] = diagram

        # Render diagram
        image = self.renderer.render_diagram(diagram)

        # Insert into document
        cursor = self.editor.textCursor()

        image_format = QTextImageFormat()
        image_format.setName(diagram_id)

        self.editor.document().addResource(
            QTextCursor.ImageResource,
            diagram_id,
            image
        )

        cursor.insertImage(image_format)
