from enum import Enum, auto
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any, Tuple, Union
from PySide6.QtCore import Qt, QPointF, QRectF, QSizeF, QObject, Signal, QPoint
from PySide6.QtGui import (
    QPainter, QPen, QBrush, QColor, QPainterPath, QFont, QFontMetrics,
    QTextFormat, QTextCharFormat, QTextFrameFormat, QTextCursor,
    QTextDocument, QTextFrame, QTextObject, QTextBlockFormat, QTextBlock,
    QMouseEvent
)
from PySide6.QtWidgets import (
    QGraphicsItem, QGraphicsRectItem, QGraphicsEllipseItem, QGraphicsPathItem,
    QGraphicsTextItem, QStyleOptionGraphicsItem, QWidget, QDialog, QVBoxLayout,
    QHBoxLayout, QLabel, QComboBox, QSpinBox, QDoubleSpinBox, QColorDialog,
    QPushButton, QDialogButtonBox, QFormLayout, QCheckBox, QGroupBox, QLineEdit,
    QGraphicsView, QGraphicsScene, QFontComboBox
)

class ShapeType(Enum):
    RECTANGLE = auto()
    ELLIPSE = auto()
    LINE = auto()
    ARROW = auto()
    TEXT_BOX = auto()
    CALL_OUT = auto()
    STAR = auto()
    CLOUD = auto()

@dataclass
class ShapeStyle:
    fill_color: QColor = field(default_factory=lambda: QColor(255, 255, 255, 200))  # White with some transparency
    stroke_color: QColor = field(default_factory=lambda: QColor(0, 0, 0))  # Black
    stroke_width: float = 1.0
    stroke_style: Qt.PenStyle = Qt.PenStyle.SolidLine
    corner_radius: float = 0.0  # For rectangles
    opacity: float = 1.0
    shadow_enabled: bool = False
    shadow_color: QColor = field(default_factory=lambda: QColor(100, 100, 100, 100))
    shadow_offset: QPointF = field(default_factory=lambda: QPointF(3, 3))
    shadow_blur: float = 5.0
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'fill_color': self.fill_color.name(),
            'fill_alpha': self.fill_color.alpha(),
            'stroke_color': self.stroke_color.name(),
            'stroke_alpha': self.stroke_color.alpha(),
            'stroke_width': self.stroke_width,
            'stroke_style': self.stroke_style,
            'corner_radius': self.corner_radius,
            'opacity': self.opacity,
            'shadow_enabled': self.shadow_enabled,
            'shadow_color': self.shadow_color.name(),
            'shadow_alpha': self.shadow_color.alpha(),
            'shadow_offset': (self.shadow_offset.x(), self.shadow_offset.y()),
            'shadow_blur': self.shadow_blur
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ShapeStyle':
        style = cls()
        
        # Fill color
        fill_color = QColor(data.get('fill_color', '#ffffff'))
        fill_color.setAlpha(data.get('fill_alpha', 200))
        style.fill_color = fill_color
        
        # Stroke color
        stroke_color = QColor(data.get('stroke_color', '#000000'))
        stroke_color.setAlpha(data.get('stroke_alpha', 255))
        style.stroke_color = stroke_color
        
        # Other properties
        style.stroke_width = data.get('stroke_width', 1.0)
        style.stroke_style = data.get('stroke_style', Qt.PenStyle.SolidLine)
        style.corner_radius = data.get('corner_radius', 0.0)
        style.opacity = data.get('opacity', 1.0)
        style.shadow_enabled = data.get('shadow_enabled', False)
        
        # Shadow
        shadow_color = QColor(data.get('shadow_color', '#646464'))
        shadow_color.setAlpha(data.get('shadow_alpha', 100))
        style.shadow_color = shadow_color
        
        shadow_offset = data.get('shadow_offset', (3, 3))
        style.shadow_offset = QPointF(*shadow_offset)
        style.shadow_blur = data.get('shadow_blur', 5.0)
        
        return style

@dataclass
class TextStyle(ShapeStyle):
    font_family: str = "Arial"
    font_size: float = 12.0
    font_bold: bool = False
    font_italic: bool = False
    font_underline: bool = False
    text_color: QColor = field(default_factory=lambda: QColor(0, 0, 0))  # Black
    text_align: Qt.Alignment = Qt.AlignLeft | Qt.AlignTop
    line_spacing: float = 1.0
    
    def to_dict(self) -> Dict[str, Any]:
        base_dict = super().to_dict()
        base_dict.update({
            'font_family': self.font_family,
            'font_size': self.font_size,
            'font_bold': self.font_bold,
            'font_italic': self.font_italic,
            'font_underline': self.font_underline,
            'text_color': self.text_color.name(),
            'text_alpha': self.text_color.alpha(),
            'text_align': int(self.text_align),
            'line_spacing': self.line_spacing
        })
        return base_dict
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'TextStyle':
        style = super().from_dict(data)
        style.font_family = data.get('font_family', 'Arial')
        style.font_size = data.get('font_size', 12.0)
        style.font_bold = data.get('font_bold', False)
        style.font_italic = data.get('font_italic', False)
        style.font_underline = data.get('font_underline', False)
        
        text_color = QColor(data.get('text_color', '#000000'))
        text_color.setAlpha(data.get('text_alpha', 255))
        style.text_color = text_color
        
        style.text_align = Qt.Alignment(data.get('text_align', int(Qt.AlignLeft | Qt.AlignTop)))
        style.line_spacing = data.get('line_spacing', 1.0)
        
        return style

@dataclass
class ShapeProperties:
    position: QPointF = field(default_factory=lambda: QPointF(0, 0))
    size: QSizeF = field(default_factory=lambda: QSizeF(100, 100))
    rotation: float = 0.0
    z_value: int = 0
    locked: bool = False
    visible: bool = True
    name: str = "Shape"
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'x': self.position.x(),
            'y': self.position.y(),
            'width': self.size.width(),
            'height': self.size.height(),
            'rotation': self.rotation,
            'z_value': self.z_value,
            'locked': self.locked,
            'visible': self.visible,
            'name': self.name
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ShapeProperties':
        props = cls()
        props.position = QPointF(data.get('x', 0), data.get('y', 0))
        props.size = QSizeF(data.get('width', 100), data.get('height', 100))
        props.rotation = data.get('rotation', 0.0)
        props.z_value = data.get('z_value', 0)
        props.locked = data.get('locked', False)
        props.visible = data.get('visible', True)
        props.name = data.get('name', 'Shape')
        return props

class Shape(QGraphicsItem):
    """Base class for all shapes."""
    def __init__(self, shape_type: ShapeType, parent=None):
        super().__init__(parent)
        self.shape_type = shape_type
        self.properties = ShapeProperties()
        self.style = ShapeStyle()
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsFocusable)
        self.setAcceptHoverEvents(True)
    
    def boundingRect(self) -> QRectF:
        return QRectF(QPointF(0, 0), self.properties.size)
    
    def paint(self, painter: QPainter, option: QStyleOptionGraphicsItem, widget: QWidget = None):
        # Save painter state
        painter.save()
        
        # Apply shape style
        pen = QPen(
            self.style.stroke_color,
            self.style.stroke_width,
            self.style.stroke_style
        )
        brush = QBrush(self.style.fill_color)
        
        painter.setPen(pen)
        painter.setBrush(brush)
        painter.setOpacity(self.style.opacity)
        
        # Draw shape (implemented by subclasses)
        self.draw_shape(painter)
        
        # Draw selection handles if selected
        if self.isSelected():
            self.draw_selection_handles(painter)
        
        # Restore painter state
        painter.restore()
    
    def draw_shape(self, painter: QPainter):
        """Draw the shape. Must be implemented by subclasses."""
        pass
    
    def draw_selection_handles(self, painter: QPainter):
        """Draw selection handles around the shape."""
        rect = self.boundingRect()
        pen = QPen(Qt.GlobalColor.blue, 1, Qt.PenStyle.DashLine)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRect(rect)
        
        # Draw resize handles
        handle_size = 6
        half_handle = handle_size / 2
        
        # Corner handles
        for x in [rect.left(), rect.right()]:
            for y in [rect.top(), rect.bottom()]:
                painter.fillRect(
                    x - half_handle, 
                    y - half_handle, 
                    handle_size, 
                    handle_size, 
                    Qt.GlobalColor.blue
                )
        
        # Edge handles
        for x in [rect.left(), rect.center().x(), rect.right()]:
            for y in [rect.top(), rect.center().y(), rect.bottom()]:
                if (x == rect.left() or x == rect.right()) and \
                   (y == rect.top() or y == rect.bottom()):
                    continue  # Skip corners (already drawn)
                if x == rect.center().x() and y == rect.center().y():
                    continue  # Skip center
                    
                painter.fillRect(
                    x - half_handle, 
                    y - half_handle, 
                    handle_size, 
                    handle_size, 
                    Qt.GlobalColor.blue
                )
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert shape to a dictionary for serialization."""
        return {
            'type': self.shape_type.name,
            'properties': self.properties.to_dict(),
            'style': self.style.to_dict()
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Shape':
        """Create a shape from a dictionary."""
        shape_type = ShapeType[data.get('type', 'RECTANGLE')]
        shape = cls.create_shape(shape_type)
        
        if 'properties' in data:
            shape.properties = ShapeProperties.from_dict(data['properties'])
        
        if 'style' in data:
            shape.style = ShapeStyle.from_dict(data['style'])
        
        return shape
    
    @classmethod
    def create_shape(cls, shape_type: ShapeType) -> 'Shape':
        """Factory method to create a shape of the specified type."""
        if shape_type == ShapeType.RECTANGLE:
            return RectangleShape()
        elif shape_type == ShapeType.ELLIPSE:
            return EllipseShape()
        elif shape_type == ShapeType.LINE:
            return LineShape()
        elif shape_type == ShapeType.ARROW:
            return ArrowShape()
        elif shape_type == ShapeType.TEXT_BOX:
            return TextBoxShape()
        elif shape_type == ShapeType.CALL_OUT:
            return CallOutShape()
        elif shape_type == ShapeType.STAR:
            return StarShape()
        elif shape_type == ShapeType.CLOUD:
            return CloudShape()
        else:
            return RectangleShape()  # Default to rectangle

class RectangleShape(Shape):
    """A rectangle shape."""
    def __init__(self):
        super().__init__(ShapeType.RECTANGLE)
        self.properties.name = "Rectangle"
    
    def draw_shape(self, painter: QPainter):
        rect = self.boundingRect()
        if self.style.corner_radius > 0:
            painter.drawRoundedRect(rect, self.style.corner_radius, self.style.corner_radius)
        else:
            painter.drawRect(rect)

class EllipseShape(Shape):
    """An ellipse shape."""
    def __init__(self):
        super().__init__(ShapeType.ELLIPSE)
        self.properties.name = "Ellipse"
    
    def draw_shape(self, painter: QPainter):
        painter.drawEllipse(self.boundingRect())

class LineShape(Shape):
    """A line shape."""
    def __init__(self):
        super().__init__(ShapeType.LINE)
        self.properties.name = "Line"
        self.start_point = QPointF(0, 0)
        self.end_point = QPointF(100, 100)
    
    def draw_shape(self, painter: QPainter):
        painter.drawLine(self.start_point, self.end_point)
    
    def boundingRect(self) -> QRectF:
        return QRectF(self.start_point, self.end_point).normalized().adjusted(-5, -5, 5, 5)

class ArrowShape(LineShape):
    """An arrow shape."""
    def __init__(self):
        super().__init__()
        self.shape_type = ShapeType.ARROW
        self.properties.name = "Arrow"
        self.arrow_size = 10
    
    def draw_shape(self, painter: QPainter):
        # Draw the line
        super().draw_shape(painter)
        
        # Draw arrow head
        line = self.end_point - self.start_point
        angle = line.angle()
        
        # Calculate arrow points
        arrow_p1 = self.end_point
        arrow_p2 = self.end_point - QPointF(
            self.arrow_size * 0.5 * (1 + 0.5 * (angle % 90) / 90),
            self.arrow_size * 0.5 * (1 - 0.5 * (angle % 90) / 90)
        )
        arrow_p3 = self.end_point - QPointF(
            -self.arrow_size * 0.5 * (1 - 0.5 * (angle % 90) / 90),
            self.arrow_size * 0.5 * (1 + 0.5 * (angle % 90) / 90)
        )
        
        # Draw arrow head
        path = QPainterPath()
        path.moveTo(arrow_p1)
        path.lineTo(arrow_p2)
        path.lineTo(arrow_p3)
        path.closeSubpath()
        
        painter.setBrush(QBrush(self.style.stroke_color))
        painter.drawPath(path)

class TextBoxShape(Shape):
    """A text box shape."""
    def __init__(self):
        super().__init__(ShapeType.TEXT_BOX)
        self.properties.name = "Text Box"
        self.text = "Double-click to edit"
        self.text_style = TextStyle()
        self.text_style.fill_color = QColor(255, 255, 255, 200)  # Semi-transparent white
        self.text_style.stroke_color = QColor(0, 0, 0)  # Black border
        self.text_style.stroke_width = 1.0
        self.text_style.corner_radius = 5.0
        self.text_style.font_size = 12.0
    
    def draw_shape(self, painter: QPainter):
        # Draw background
        rect = self.boundingRect()
        if self.style.corner_radius > 0:
            path = QPainterPath()
            path.addRoundedRect(rect, self.style.corner_radius, self.style.corner_radius)
            painter.fillPath(path, QBrush(self.style.fill_color))
            painter.drawPath(path)
        else:
            painter.fillRect(rect, QBrush(self.style.fill_color))
            painter.drawRect(rect)
        
        # Draw text
        text_rect = rect.adjusted(5, 5, -5, -5)  # Add padding
        painter.setPen(QPen(self.text_style.text_color))
        
        # Set font
        font = QFont(self.text_style.font_family, int(self.text_style.font_size))
        font.setBold(self.text_style.font_bold)
        font.setItalic(self.text_style.font_italic)
        font.setUnderline(self.text_style.font_underline)
        painter.setFont(font)
        
        # Draw text with word wrap
        painter.drawText(text_rect, int(self.text_style.text_align), self.text)
    
    def to_dict(self) -> Dict[str, Any]:
        data = super().to_dict()
        data.update({
            'text': self.text,
            'text_style': self.text_style.to_dict()
        })
        return data
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'TextBoxShape':
        shape = super().from_dict(data)
        if isinstance(shape, TextBoxShape):
            shape.text = data.get('text', '')
            if 'text_style' in data:
                shape.text_style = TextStyle.from_dict(data['text_style'])
        return shape

class CallOutShape(TextBoxShape):
    """A callout shape with a pointer."""
    def __init__(self):
        super().__init__()
        self.shape_type = ShapeType.CALL_OUT
        self.properties.name = "Call Out"
        self.pointer_pos = QPointF(50, 100)  # Relative to shape
    
    def draw_shape(self, painter: QPainter):
        # Draw the main rectangle (from parent)
        rect = self.boundingRect()
        
        # Draw the callout bubble
        path = QPainterPath()
        path.addRoundedRect(rect, self.style.corner_radius, self.style.corner_radius)
        
        # Add pointer
        pointer_start = QPointF(rect.center().x(), rect.bottom())
        pointer_tip = self.pointer_pos
        
        path.moveTo(pointer_start)
        path.lineTo(pointer_tip)
        path.lineTo(pointer_start + QPointF(20, 0))
        path.closeSubpath()
        
        # Fill and stroke
        painter.fillPath(path, QBrush(self.style.fill_color))
        painter.strokePath(path, QPen(self.style.stroke_color, self.style.stroke_width))
        
        # Draw text
        text_rect = rect.adjusted(5, 5, -5, -5)  # Add padding
        painter.setPen(QPen(self.text_style.text_color))
        
        # Set font
        font = QFont(self.text_style.font_family, int(self.text_style.font_size))
        font.setBold(self.text_style.font_bold)
        font.setItalic(self.text_style.font_italic)
        font.setUnderline(self.text_style.font_underline)
        painter.setFont(font)
        
        # Draw text with word wrap
        painter.drawText(text_rect, int(self.text_style.text_align), self.text)

class StarShape(Shape):
    """A star shape."""
    def __init__(self):
        super().__init__(ShapeType.STAR)
        self.properties.name = "Star"
        self.points = 5
        self.inner_radius = 0.4  # As a fraction of outer radius
    
    def draw_shape(self, painter: QPainter):
        import math
        rect = self.boundingRect()
        center = rect.center()
        radius = min(rect.width(), rect.height()) * 0.5
        inner_radius = radius * self.inner_radius
        
        path = QPainterPath()
        
        for i in range(self.points * 2):
            angle = (i * 2 * 3.14159) / (self.points * 2) - 3.14159 / 2
            r = radius if i % 2 == 0 else inner_radius
            x = center.x() + r * math.cos(angle)
            y = center.y() + r * math.sin(angle)
            
            if i == 0:
                path.moveTo(x, y)
            else:
                path.lineTo(x, y)
        
        path.closeSubpath()
        painter.drawPath(path)

class CloudShape(Shape):
    """A cloud shape."""
    def __init__(self):
        super().__init__(ShapeType.CLOUD)
        self.properties.name = "Cloud"
    
    def draw_shape(self, painter: QPainter):
        rect = self.boundingRect()
        center = rect.center()
        radius = min(rect.width(), rect.height()) * 0.5
        
        path = QPainterPath()
        
        # Draw cloud shape using bezier curves
        path.moveTo(center.x() - radius * 0.7, center.y() - radius * 0.3)
        
        # Top curve
        path.cubicTo(
            center.x() - radius * 0.3, center.y() - radius * 0.7,
            center.x() + radius * 0.3, center.y() - radius * 0.7,
            center.x() + radius * 0.7, center.y() - radius * 0.3
        )
        
        # Right curve
        path.cubicTo(
            center.x() + radius * 0.9, center.y(),
            center.x() + radius * 0.7, center.y() + radius * 0.5,
            center.x(), center.y() + radius * 0.5
        )
        
        # Bottom curve
        path.cubicTo(
            center.x() - radius * 0.7, center.y() + radius * 0.5,
            center.x() - radius * 0.9, center.y(),
            center.x() - radius * 0.7, center.y() - radius * 0.3
        )
        
        painter.drawPath(path)

class ShapeManager(QObject):
    """Manages shapes in a document."""
    shapeAdded = Signal(object)  # Shape
    shapeRemoved = Signal(object)  # Shape
    shapeChanged = Signal(object)  # Shape
    selectionChanged = Signal(list)  # List of selected shapes
    
    def __init__(self, document: QTextDocument, parent=None):
        super().__init__(parent)
        self.document = document
        self.shapes: List[Shape] = []
        self.selected_shapes: List[Shape] = []
        self._next_z = 1
    
    def add_shape(self, shape: Shape) -> Shape:
        """Add a shape to the document."""
        shape.setZValue(self._next_z)
        self._next_z += 1
        self.shapes.append(shape)
        self.shapeAdded.emit(shape)
        return shape
    
    def remove_shape(self, shape: Shape):
        """Remove a shape from the document."""
        if shape in self.shapes:
            self.shapes.remove(shape)
            self.deselect_shape(shape)
            self.shapeRemoved.emit(shape)
    
    def select_shape(self, shape: Shape, multi_select: bool = False):
        """Select a shape."""
        if not multi_select:
            self.clear_selection()
        
        if shape not in self.selected_shapes:
            self.selected_shapes.append(shape)
            shape.setSelected(True)
            self.selectionChanged.emit(self.selected_shapes)
    
    def deselect_shape(self, shape: Shape):
        """Deselect a shape."""
        if shape in self.selected_shapes:
            self.selected_shapes.remove(shape)
            shape.setSelected(False)
            self.selectionChanged.emit(self.selected_shapes)
    
    def clear_selection(self):
        """Clear the current selection."""
        for shape in self.selected_shapes:
            shape.setSelected(False)
        self.selected_shapes.clear()
        self.selectionChanged.emit([])
    
    def bring_to_front(self, shape: Shape):
        """Bring a shape to the front."""
        if shape in self.shapes:
            max_z = max((s.zValue() for s in self.shapes), default=0)
            shape.setZValue(max_z + 1)
            self.shapeChanged.emit(shape)
    
    def send_to_back(self, shape: Shape):
        """Send a shape to the back."""
        if shape in self.shapes:
            min_z = min((s.zValue() for s in self.shapes), default=0)
            shape.setZValue(min_z - 1)
            self.shapeChanged.emit(shape)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert all shapes to a dictionary for serialization."""
        return {
            'shapes': [shape.to_dict() for shape in self.shapes],
            'next_z': self._next_z
        }
    
    def from_dict(self, data: Dict[str, Any]):
        """Load shapes from a dictionary."""
        self.clear_selection()
        self.shapes.clear()
        
        for shape_data in data.get('shapes', []):
            shape = Shape.from_dict(shape_data)
            if shape:
                self.shapes.append(shape)
        
        self._next_z = data.get('next_z', len(self.shapes) + 1)
        
        # Notify about the change
        for shape in self.shapes:
            self.shapeAdded.emit(shape)

class ShapeToolbar(QWidget):
    """Toolbar for working with shapes."""
    def __init__(self, shape_manager: ShapeManager, parent=None):
        super().__init__(parent)
        self.shape_manager = shape_manager
        self.init_ui()
    
    def init_ui(self):
        layout = QHBoxLayout()
        
        # Shape type selector
        self.shape_combo = QComboBox()
        for shape_type in ShapeType:
            self.shape_combo.addItem(shape_type.name.replace('_', ' ').title(), shape_type)
        
        # Add shape button
        add_btn = QPushButton("Add Shape")
        add_btn.clicked.connect(self._on_add_shape)
        
        # Bring to front/send to back
        front_btn = QPushButton("Front")
        front_btn.clicked.connect(self._on_bring_to_front)
        
        back_btn = QPushButton("Back")
        back_btn.clicked.connect(self._on_send_to_back)
        
        # Delete button
        delete_btn = QPushButton("Delete")
        delete_btn.clicked.connect(self._on_delete_shape)
        
        # Add widgets to layout
        layout.addWidget(QLabel("Shape:"))
        layout.addWidget(self.shape_combo)
        layout.addWidget(add_btn)
        layout.addWidget(front_btn)
        layout.addWidget(back_btn)
        layout.addWidget(delete_btn)
        layout.addStretch()
        
        self.setLayout(layout)
    
    def _on_add_shape(self):
        shape_type = self.shape_combo.currentData()
        shape = Shape.create_shape(shape_type)
        self.shape_manager.add_shape(shape)
    
    def _on_bring_to_front(self):
        for shape in self.shape_manager.selected_shapes[:]:
            self.shape_manager.bring_to_front(shape)
    
    def _on_send_to_back(self):
        for shape in self.shape_manager.selected_shapes[:]:
            self.shape_manager.send_to_back(shape)
    
    def _on_delete_shape(self):
        for shape in self.shape_manager.selected_shapes[:]:
            self.shape_manager.remove_shape(shape)

class ShapePropertiesDialog(QDialog):
    """Dialog for editing shape properties."""
    def __init__(self, shape: Shape, parent=None):
        super().__init__(parent)
        self.shape = shape
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("Shape Properties")
        self.setMinimumSize(400, 500)
        
        layout = QVBoxLayout()
        
        # Name
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Name:"))
        self.name_edit = QLineEdit(self.shape.properties.name)
        name_layout.addWidget(self.name_edit)
        layout.addLayout(name_layout)
        
        # Position and size
        pos_size_group = QGroupBox("Position & Size")
        pos_size_layout = QFormLayout()
        
        self.x_spin = QDoubleSpinBox()
        self.x_spin.setRange(-10000, 10000)
        self.x_spin.setValue(self.shape.properties.position.x())
        pos_size_layout.addRow("X:", self.x_spin)
        
        self.y_spin = QDoubleSpinBox()
        self.y_spin.setRange(-10000, 10000)
        self.y_spin.setValue(self.shape.properties.position.y())
        pos_size_layout.addRow("Y:", self.y_spin)
        
        self.width_spin = QDoubleSpinBox()
        self.width_spin.setRange(1, 10000)
        self.width_spin.setValue(self.shape.properties.size.width())
        pos_size_layout.addRow("Width:", self.width_spin)
        
        self.height_spin = QDoubleSpinBox()
        self.height_spin.setRange(1, 10000)
        self.height_spin.setValue(self.shape.properties.size.height())
        pos_size_layout.addRow("Height:", self.height_spin)
        
        self.rotation_spin = QDoubleSpinBox()
        self.rotation_spin.setRange(0, 359)
        self.rotation_spin.setValue(self.shape.properties.rotation)
        self.rotation_spin.setSuffix("°")
        pos_size_layout.addRow("Rotation:", self.rotation_spin)
        
        pos_size_group.setLayout(pos_size_layout)
        layout.addWidget(pos_size_group)
        
        # Style
        style_group = QGroupBox("Style")
        style_layout = QFormLayout()
        
        # Fill color
        self.fill_color_btn = QPushButton()
        self.fill_color_btn.setStyleSheet(
            f"background-color: {self.shape.style.fill_color.name()};"
            f"border: 1px solid #000000;"
        )
        self.fill_color_btn.clicked.connect(self._on_choose_fill_color)
        style_layout.addRow("Fill Color:", self.fill_color_btn)
        
        # Stroke color
        self.stroke_color_btn = QPushButton()
        self.stroke_color_btn.setStyleSheet(
            f"background-color: {self.shape.style.stroke_color.name()};"
            f"border: 1px solid #000000;"
        )
        self.stroke_color_btn.clicked.connect(self._on_choose_stroke_color)
        style_layout.addRow("Stroke Color:", self.stroke_color_btn)
        
        # Stroke width
        self.stroke_width_spin = QDoubleSpinBox()
        self.stroke_width_spin.setRange(0.1, 100)
        self.stroke_width_spin.setValue(self.shape.style.stroke_width)
        style_layout.addRow("Stroke Width:", self.stroke_width_spin)
        
        # Stroke style
        self.stroke_style_combo = QComboBox()
        self.stroke_style_combo.addItem("Solid", Qt.PenStyle.SolidLine)
        self.stroke_style_combo.addItem("Dash", Qt.PenStyle.DashLine)
        self.stroke_style_combo.addItem("Dot", Qt.PenStyle.DotLine)
        self.stroke_style_combo.addItem("Dash Dot", Qt.PenStyle.DashDotLine)
        self.stroke_style_combo.addItem("Dash Dot Dot", Qt.PenStyle.DashDotDotLine)
        
        # Find the current stroke style
        index = self.stroke_style_combo.findData(self.shape.style.stroke_style)
        if index >= 0:
            self.stroke_style_combo.setCurrentIndex(index)
            
        style_layout.addRow("Stroke Style:", self.stroke_style_combo)
        
        # Corner radius (for rectangles)
        if hasattr(self.shape, 'style') and hasattr(self.shape.style, 'corner_radius'):
            self.corner_radius_spin = QDoubleSpinBox()
            self.corner_radius_spin.setRange(0, 1000)
            self.corner_radius_spin.setValue(self.shape.style.corner_radius)
            style_layout.addRow("Corner Radius:", self.corner_radius_spin)
        
        style_group.setLayout(style_layout)
        layout.addWidget(style_group)
        
        # Text properties (for text boxes)
        if isinstance(self.shape, (TextBoxShape, CallOutShape)):
            text_group = QGroupBox("Text")
            text_layout = QFormLayout()
            
            self.text_edit = QTextEdit()
            self.text_edit.setPlainText(self.shape.text)
            text_layout.addRow("Text:", self.text_edit)
            
            # Font family
            self.font_family_combo = QFontComboBox()
            self.font_family_combo.setCurrentFont(QFont(self.shape.text_style.font_family))
            text_layout.addRow("Font:", self.font_family_combo)
            
            # Font size
            self.font_size_spin = QDoubleSpinBox()
            self.font_size_spin.setRange(1, 500)
            self.font_size_spin.setValue(self.shape.text_style.font_size)
            text_layout.addRow("Size:", self.font_size_spin)
            
            # Font style
            self.bold_check = QCheckBox("Bold")
            self.bold_check.setChecked(self.shape.text_style.font_bold)
            
            self.italic_check = QCheckBox("Italic")
            self.italic_check.setChecked(self.shape.text_style.font_italic)
            
            self.underline_check = QCheckBox("Underline")
            self.underline_check.setChecked(self.shape.text_style.font_underline)
            
            font_style_layout = QHBoxLayout()
            font_style_layout.addWidget(self.bold_check)
            font_style_layout.addWidget(self.italic_check)
            font_style_layout.addWidget(self.underline_check)
            text_layout.addRow("Style:", font_style_layout)
            
            # Text color
            self.text_color_btn = QPushButton()
            self.text_color_btn.setStyleSheet(
                f"background-color: {self.shape.text_style.text_color.name()};"
                f"border: 1px solid #000000;"
            )
            self.text_color_btn.clicked.connect(self._on_choose_text_color)
            text_layout.addRow("Text Color:", self.text_color_btn)
            
            # Text alignment
            self.align_combo = QComboBox()
            self.align_combo.addItem("Left", Qt.AlignLeft)
            self.align_combo.addItem("Center", Qt.AlignCenter)
            self.align_combo.addItem("Right", Qt.AlignRight)
            
            # Find the current alignment
            index = self.align_combo.findData(self.shape.text_style.text_align & (Qt.AlignLeft | Qt.AlignHCenter | Qt.AlignRight))
            if index >= 0:
                self.align_combo.setCurrentIndex(index)
                
            text_layout.addRow("Alignment:", self.align_combo)
            
            text_group.setLayout(text_layout)
            layout.addWidget(text_group)
        
        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def _on_choose_fill_color(self):
        color = QColorDialog.getColor(self.shape.style.fill_color, self, "Choose Fill Color")
        if color.isValid():
            self.shape.style.fill_color = color
            self.fill_color_btn.setStyleSheet(
                f"background-color: {color.name()};"
                f"border: 1px solid #000000;"
            )
    
    def _on_choose_stroke_color(self):
        color = QColorDialog.getColor(self.shape.style.stroke_color, self, "Choose Stroke Color")
        if color.isValid():
            self.shape.style.stroke_color = color
            self.stroke_color_btn.setStyleSheet(
                f"background-color: {color.name()};"
                f"border: 1px solid #000000;"
            )
    
    def _on_choose_text_color(self):
        if not hasattr(self, 'text_color_btn'):
            return
            
        color = QColorDialog.getColor(self.shape.text_style.text_color, self, "Choose Text Color")
        if color.isValid():
            self.shape.text_style.text_color = color
            self.text_color_btn.setStyleSheet(
                f"background-color: {color.name()};"
                f"border: 1px solid #000000;"
            )
    
    def accept(self):
        # Update shape properties
        self.shape.properties.name = self.name_edit.text()
        self.shape.properties.position = QPointF(self.x_spin.value(), self.y_spin.value())
        self.shape.properties.size = QSizeF(self.width_spin.value(), self.height_spin.value())
        self.shape.properties.rotation = self.rotation_spin.value()
        
        # Update style
        self.shape.style.stroke_width = self.stroke_width_spin.value()
        self.shape.style.stroke_style = self.stroke_style_combo.currentData()
        
        # Update corner radius if applicable
        if hasattr(self, 'corner_radius_spin'):
            self.shape.style.corner_radius = self.corner_radius_spin.value()
        
        # Update text properties if applicable
        if isinstance(self.shape, (TextBoxShape, CallOutShape)):
            self.shape.text = self.text_edit.toPlainText()
            self.shape.text_style.font_family = self.font_family_combo.currentFont().family()
            self.shape.text_style.font_size = self.font_size_spin.value()
            self.shape.text_style.font_bold = self.bold_check.isChecked()
            self.shape.text_style.font_italic = self.italic_check.isChecked()
            self.shape.text_style.font_underline = self.underline_check.isChecked()
            self.shape.text_style.text_align = self.align_combo.currentData()
        
        super().accept()

class ShapeView(QGraphicsView):
    """View for editing shapes."""
    def __init__(self, shape_manager: ShapeManager, parent=None):
        super().__init__(parent)
        self.shape_manager = shape_manager
        self.setScene(QGraphicsScene(self))
        self.setRenderHint(QPainter.Antialiasing)
        self.setViewportUpdateMode(QGraphicsView.FullViewportUpdate)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        self.setDragMode(QGraphicsView.RubberBandDrag)
        
        # Connect signals
        self.scene().selectionChanged.connect(self._on_selection_changed)
        
        # Add shapes to scene
        for shape in self.shape_manager.shapes:
            self.scene().addItem(shape)
    
    def _on_selection_changed(self):
        selected_shapes = [item for item in self.scene().selectedItems() 
                          if isinstance(item, Shape)]
        
        # Update shape manager's selection
        self.shape_manager.clear_selection()
        for shape in selected_shapes:
            self.shape_manager.select_shape(shape, multi_select=True)
    
    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            # Handle shape creation or selection
            pass
        super().mousePressEvent(event)
    
    def mouseDoubleClickEvent(self, event: QMouseEvent):
        item = self.itemAt(event.pos())
        if isinstance(item, Shape):
            # Open properties dialog
            dialog = ShapePropertiesDialog(item, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                item.update()
        super().mouseDoubleClickEvent(event)

class ShapeEditor(QWidget):
    """Main widget for shape editing."""
    def __init__(self, document: QTextDocument, parent=None):
        super().__init__(parent)
        self.document = document
        self.shape_manager = ShapeManager(document)
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Create toolbar
        self.toolbar = ShapeToolbar(self.shape_manager, self)
        
        # Create view
        self.view = ShapeView(self.shape_manager, self)
        
        # Add widgets to layout
        layout.addWidget(self.toolbar)
        layout.addWidget(self.view)
        
        self.setLayout(layout)
    
    def add_shape(self, shape_type: ShapeType):
        """Add a new shape to the editor."""
        shape = Shape.create_shape(shape_type)
        self.shape_manager.add_shape(shape)
        self.view.scene().addItem(shape)
        return shape
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert editor state to a dictionary."""
        return self.shape_manager.to_dict()
    
    def from_dict(self, data: Dict[str, Any]):
        """Load editor state from a dictionary."""
        self.shape_manager.from_dict(data)
        
        # Clear existing items
        self.view.scene().clear()
        
        # Add shapes to scene
        for shape in self.shape_manager.shapes:
            self.view.scene().addItem(shape)

# Example usage
if __name__ == "__main__":
    import sys
    from PySide6.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    
    # Create a document
    doc = QTextDocument()
    
    # Create the editor
    editor = ShapeEditor(doc)
    editor.resize(800, 600)
    editor.setWindowTitle("Shape Editor")
    editor.show()
    
    # Add some example shapes
    rect = editor.add_shape(ShapeType.RECTANGLE)
    rect.properties.position = QPointF(100, 100)
    rect.properties.size = QSizeF(200, 150)
    rect.style.fill_color = QColor(255, 200, 200, 200)
    rect.style.stroke_color = QColor(255, 0, 0)
    rect.style.stroke_width = 2.0
    
    circle = editor.add_shape(ShapeType.ELLIPSE)
    circle.properties.position = QPointF(150, 150)
    circle.properties.size = QSizeF(150, 150)
    circle.style.fill_color = QColor(200, 200, 255, 200)
    circle.style.stroke_color = QColor(0, 0, 255)
    circle.style.stroke_width = 2.0
    
    text_box = editor.add_shape(ShapeType.TEXT_BOX)
    text_box.properties.position = QPointF(200, 200)
    text_box.properties.size = QSizeF(150, 100)
    text_box.text = "Double-click to edit\nThis is a text box"
    text_box.text_style.font_size = 14
    text_box.text_style.font_bold = True
    text_box.text_style.text_align = Qt.AlignCenter
    
    sys.exit(app.exec())
