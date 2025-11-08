"""Rulers for PyWord - horizontal and vertical rulers with indent markers."""

from PySide6.QtWidgets import QWidget
from PySide6.QtCore import Qt, QSize, QPoint
from PySide6.QtGui import QPainter, QColor, QPen, QFont, QPainterPath


class HorizontalRuler(QWidget):
    """Horizontal ruler widget showing measurements and indent markers."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(25)
        self.unit = "in"  # inches
        self.dpi = 96  # dots per inch
        self.zoom = 1.0
        self.left_indent = 0
        self.right_indent = 0
        self.first_line_indent = 0

        # Microsoft Word colors
        self.bg_color = QColor("#FFFFFF")
        self.border_color = QColor("#D2D0CE")
        self.text_color = QColor("#605E5C")
        self.marker_color = QColor("#0078D4")

    def paintEvent(self, event):
        """Paint the horizontal ruler."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Draw background
        painter.fillRect(self.rect(), self.bg_color)

        # Draw bottom border
        painter.setPen(QPen(self.border_color, 1))
        painter.drawLine(0, self.height() - 1, self.width(), self.height() - 1)

        # Draw ruler markings
        self.draw_ruler_markings(painter)

        # Draw indent markers
        self.draw_indent_markers(painter)

    def draw_ruler_markings(self, painter):
        """Draw the ruler tick marks and numbers."""
        painter.setPen(QPen(self.text_color, 1))

        # Set font for numbers
        font = QFont("Segoe UI", 8)
        painter.setFont(font)

        # Calculate pixels per inch
        pixels_per_inch = self.dpi * self.zoom

        # Draw tick marks every 1/8 inch
        for i in range(0, int(self.width() / (pixels_per_inch / 8))):
            x = int(i * pixels_per_inch / 8)
            if x > self.width():
                break

            # Every inch: long tick + number
            if i % 8 == 0:
                painter.drawLine(x, self.height() - 10, x, self.height() - 1)
                # Draw inch number
                inch = i // 8
                painter.drawText(x + 2, 12, str(inch))
            # Every 1/2 inch: medium tick
            elif i % 4 == 0:
                painter.drawLine(x, self.height() - 7, x, self.height() - 1)
            # Every 1/4 inch: small tick
            elif i % 2 == 0:
                painter.drawLine(x, self.height() - 5, x, self.height() - 1)
            # Every 1/8 inch: tiny tick
            else:
                painter.drawLine(x, self.height() - 3, x, self.height() - 1)

    def draw_indent_markers(self, painter):
        """Draw indent markers (triangular markers for indents)."""
        painter.setPen(QPen(self.marker_color, 1))
        painter.setBrush(self.marker_color)

        pixels_per_inch = self.dpi * self.zoom

        # Left indent marker (bottom triangle)
        left_x = int(self.left_indent * pixels_per_inch)
        painter.drawPolygon([
            QPoint(left_x - 4, self.height() - 2),
            QPoint(left_x + 4, self.height() - 2),
            QPoint(left_x, self.height() - 8)
        ])

        # First line indent marker (top triangle)
        first_line_x = int(self.first_line_indent * pixels_per_inch)
        painter.drawPolygon([
            QPoint(first_line_x - 4, 2),
            QPoint(first_line_x + 4, 2),
            QPoint(first_line_x, 8)
        ])

    def set_zoom(self, zoom):
        """Set the zoom level."""
        self.zoom = zoom / 100.0
        self.update()

    def set_left_indent(self, inches):
        """Set the left indent in inches."""
        self.left_indent = inches
        self.update()

    def set_first_line_indent(self, inches):
        """Set the first line indent in inches."""
        self.first_line_indent = inches
        self.update()


class VerticalRuler(QWidget):
    """Vertical ruler widget showing measurements."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(25)
        self.unit = "in"  # inches
        self.dpi = 96  # dots per inch
        self.zoom = 1.0

        # Microsoft Word colors
        self.bg_color = QColor("#FFFFFF")
        self.border_color = QColor("#D2D0CE")
        self.text_color = QColor("#605E5C")

    def paintEvent(self, event):
        """Paint the vertical ruler."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Draw background
        painter.fillRect(self.rect(), self.bg_color)

        # Draw right border
        painter.setPen(QPen(self.border_color, 1))
        painter.drawLine(self.width() - 1, 0, self.width() - 1, self.height())

        # Draw ruler markings
        self.draw_ruler_markings(painter)

    def draw_ruler_markings(self, painter):
        """Draw the ruler tick marks and numbers."""
        painter.setPen(QPen(self.text_color, 1))

        # Set font for numbers
        font = QFont("Segoe UI", 8)
        painter.setFont(font)

        # Calculate pixels per inch
        pixels_per_inch = self.dpi * self.zoom

        # Draw tick marks every 1/8 inch
        for i in range(0, int(self.height() / (pixels_per_inch / 8))):
            y = int(i * pixels_per_inch / 8)
            if y > self.height():
                break

            # Every inch: long tick + number
            if i % 8 == 0:
                painter.drawLine(self.width() - 10, y, self.width() - 1, y)
                # Draw inch number (rotated)
                inch = i // 8
                painter.save()
                painter.translate(8, y + 10)
                painter.rotate(-90)
                painter.drawText(0, 0, str(inch))
                painter.restore()
            # Every 1/2 inch: medium tick
            elif i % 4 == 0:
                painter.drawLine(self.width() - 7, y, self.width() - 1, y)
            # Every 1/4 inch: small tick
            elif i % 2 == 0:
                painter.drawLine(self.width() - 5, y, self.width() - 1, y)
            # Every 1/8 inch: tiny tick
            else:
                painter.drawLine(self.width() - 3, y, self.width() - 1, y)

    def set_zoom(self, zoom):
        """Set the zoom level."""
        self.zoom = zoom / 100.0
        self.update()
