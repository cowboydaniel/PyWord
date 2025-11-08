"""Rulers for PyWord - horizontal and vertical rulers with indent markers."""

from PySide6.QtWidgets import QWidget
from PySide6.QtCore import Qt, QSize, QPoint
from PySide6.QtGui import QPainter, QColor, QPen, QFont, QPainterPath


class HorizontalRuler(QWidget):
    """Horizontal ruler widget showing measurements and indent markers."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(25)
        self.unit = "cm"  # centimeters
        self.dpi = 96  # dots per inch
        self.zoom = 1.0
        self.left_indent = 0
        self.right_indent = 0
        self.first_line_indent = 0

        # Page dimensions (set by parent)
        self.page_width_px = 794  # A4 width in pixels at 96 DPI
        self.page_width = 21.0  # A4 width in cm

        # Page margins (in cm)
        self.left_margin = 2.54  # 1 inch = 2.54 cm left margin
        self.right_margin = 2.54  # 1 inch = 2.54 cm right margin

        # Tab stops (list of positions in cm)
        self.tab_stops = []

        # Microsoft Word colors
        self.bg_color = QColor("#FFFFFF")
        self.border_color = QColor("#D2D0CE")
        self.text_color = QColor("#605E5C")
        self.marker_color = QColor("#0078D4")
        self.margin_color = QColor("#E6E6E6")

    def paintEvent(self, event):
        """Paint the horizontal ruler."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Draw background
        painter.fillRect(self.rect(), self.bg_color)

        # Draw margin indicators (gray areas)
        self.draw_margin_indicators(painter)

        # Draw bottom border
        painter.setPen(QPen(self.border_color, 1))
        painter.drawLine(0, self.height() - 1, self.width(), self.height() - 1)

        # Draw ruler markings
        self.draw_ruler_markings(painter)

        # Draw tab stops
        self.draw_tab_stops(painter)

        # Draw indent markers
        self.draw_indent_markers(painter)

    def draw_ruler_markings(self, painter):
        """Draw the ruler tick marks and numbers."""
        painter.setPen(QPen(self.text_color, 1))

        # Set font for numbers
        font = QFont("Segoe UI", 8)
        painter.setFont(font)

        # Calculate pixels per cm (1 inch = 2.54 cm, 96 DPI)
        pixels_per_cm = (self.dpi / 2.54) * self.zoom

        # Calculate offset where page starts (ruler "0" aligns with page left edge)
        # Page is centered, so offset = (total_width - page_width) / 2
        page_offset = (self.width() - self.page_width_px) // 2

        # Calculate how many cm to show before and after 0
        cm_before_zero = int(page_offset / pixels_per_cm) + 2
        cm_after_zero = int((self.width() - page_offset) / pixels_per_cm) + 2

        # Draw tick marks every mm, including negative numbers
        for i in range(-cm_before_zero * 10, cm_after_zero * 10):
            x = page_offset + int(i * pixels_per_cm / 10)
            if x < 0 or x > self.width():
                continue

            # Every cm: long tick + number
            if i % 10 == 0:
                painter.drawLine(x, self.height() - 10, x, self.height() - 1)
                # Draw cm number (can be negative)
                cm = i // 10
                painter.drawText(x + 2, 12, str(cm))
            # Every 5mm: medium tick
            elif i % 5 == 0:
                painter.drawLine(x, self.height() - 7, x, self.height() - 1)
            # Every mm: small tick
            else:
                painter.drawLine(x, self.height() - 4, x, self.height() - 1)

    def draw_margin_indicators(self, painter):
        """Draw gray areas showing page margins."""
        painter.setPen(Qt.NoPen)
        painter.setBrush(self.margin_color)

        # Calculate pixels per cm (1 inch = 2.54 cm, 96 DPI)
        pixels_per_cm = (self.dpi / 2.54) * self.zoom

        # Calculate page offset (where page starts)
        page_offset = (self.width() - self.page_width_px) // 2

        # Left margin (gray area from page start to left margin)
        left_margin_pixels = int(self.left_margin * pixels_per_cm)
        painter.drawRect(page_offset, 0, left_margin_pixels, self.height())

        # Right margin (gray area from page_width - right_margin to page end)
        right_margin_start = page_offset + int((self.page_width - self.right_margin) * pixels_per_cm)
        right_margin_width = int(self.right_margin * pixels_per_cm)
        painter.drawRect(right_margin_start, 0, right_margin_width, self.height())

    def draw_tab_stops(self, painter):
        """Draw tab stop indicators (small L-shaped markers)."""
        painter.setPen(QPen(self.text_color, 1))
        painter.setBrush(self.text_color)

        # Calculate pixels per cm (1 inch = 2.54 cm, 96 DPI)
        pixels_per_cm = (self.dpi / 2.54) * self.zoom

        # Calculate page offset (where page starts)
        page_offset = (self.width() - self.page_width_px) // 2

        # Start from left margin
        left_margin_offset = page_offset + int(self.left_margin * pixels_per_cm)

        for tab_position in self.tab_stops:
            x = left_margin_offset + int(tab_position * pixels_per_cm)
            # Draw L-shaped tab stop marker
            painter.drawLine(x, self.height() - 6, x, self.height() - 2)
            painter.drawLine(x, self.height() - 6, x + 3, self.height() - 6)

    def draw_indent_markers(self, painter):
        """Draw indent markers (triangular markers for indents)."""
        painter.setPen(QPen(self.marker_color, 1))
        painter.setBrush(self.marker_color)

        # Calculate pixels per cm (1 inch = 2.54 cm, 96 DPI)
        pixels_per_cm = (self.dpi / 2.54) * self.zoom

        # Calculate page offset (where page starts)
        page_offset = (self.width() - self.page_width_px) // 2

        # Start from left margin
        left_margin_offset = page_offset + int(self.left_margin * pixels_per_cm)

        # Left indent marker (bottom triangle) - positioned relative to left margin
        left_x = left_margin_offset + int(self.left_indent * pixels_per_cm)
        painter.drawPolygon([
            QPoint(left_x - 4, self.height() - 2),
            QPoint(left_x + 4, self.height() - 2),
            QPoint(left_x, self.height() - 8)
        ])

        # First line indent marker (top triangle) - positioned relative to left margin
        first_line_x = left_margin_offset + int(self.first_line_indent * pixels_per_cm)
        painter.drawPolygon([
            QPoint(first_line_x - 4, 2),
            QPoint(first_line_x + 4, 2),
            QPoint(first_line_x, 8)
        ])

    def add_tab_stop(self, cm):
        """Add a tab stop at the specified position in cm."""
        self.tab_stops.append(cm)
        self.tab_stops.sort()
        self.update()

    def remove_tab_stop(self, cm):
        """Remove a tab stop at the specified position in cm."""
        if cm in self.tab_stops:
            self.tab_stops.remove(cm)
            self.update()

    def clear_tab_stops(self):
        """Clear all tab stops."""
        self.tab_stops.clear()
        self.update()

    def set_zoom(self, zoom):
        """Set the zoom level."""
        self.zoom = zoom / 100.0
        self.update()

    def set_left_indent(self, cm):
        """Set the left indent in cm."""
        self.left_indent = cm
        self.update()

    def set_first_line_indent(self, cm):
        """Set the first line indent in cm."""
        self.first_line_indent = cm
        self.update()


class VerticalRuler(QWidget):
    """Vertical ruler widget showing measurements."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(25)
        self.unit = "cm"  # centimeters
        self.dpi = 96  # dots per inch
        self.zoom = 1.0

        # Page dimensions (set by parent)
        self.page_height_px = 1122  # A4 height in pixels at 96 DPI
        self.page_height = 29.7  # A4 height in cm

        # Page margins (in cm)
        self.top_margin = 2.54  # 1 inch = 2.54 cm top margin
        self.bottom_margin = 2.54  # 1 inch = 2.54 cm bottom margin

        # Microsoft Word colors
        self.bg_color = QColor("#FFFFFF")
        self.border_color = QColor("#D2D0CE")
        self.text_color = QColor("#605E5C")
        self.margin_color = QColor("#E6E6E6")

    def paintEvent(self, event):
        """Paint the vertical ruler."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Draw background
        painter.fillRect(self.rect(), self.bg_color)

        # Draw margin indicators (gray areas)
        self.draw_margin_indicators(painter)

        # Draw right border
        painter.setPen(QPen(self.border_color, 1))
        painter.drawLine(self.width() - 1, 0, self.width() - 1, self.height())

        # Draw ruler markings
        self.draw_ruler_markings(painter)

    def draw_margin_indicators(self, painter):
        """Draw gray areas showing page margins."""
        painter.setPen(Qt.NoPen)
        painter.setBrush(self.margin_color)

        # Calculate pixels per cm (1 inch = 2.54 cm, 96 DPI)
        pixels_per_cm = (self.dpi / 2.54) * self.zoom

        # Calculate page offset (where page starts vertically)
        # Page is at top of workspace (no vertical centering)
        page_offset = 0

        # Top margin (gray area from page start to top margin)
        top_margin_pixels = int(self.top_margin * pixels_per_cm)
        painter.drawRect(0, page_offset, self.width(), top_margin_pixels)

        # Bottom margin (gray area from page_height - bottom_margin to page end)
        bottom_margin_start = page_offset + int((self.page_height - self.bottom_margin) * pixels_per_cm)
        bottom_margin_height = int(self.bottom_margin * pixels_per_cm)
        painter.drawRect(0, bottom_margin_start, self.width(), bottom_margin_height)

    def draw_ruler_markings(self, painter):
        """Draw the ruler tick marks and numbers."""
        painter.setPen(QPen(self.text_color, 1))

        # Set font for numbers
        font = QFont("Segoe UI", 8)
        painter.setFont(font)

        # Calculate pixels per cm (1 inch = 2.54 cm, 96 DPI)
        pixels_per_cm = (self.dpi / 2.54) * self.zoom

        # Calculate offset where page starts (ruler "0" aligns with page top edge)
        # Page is at top of workspace (no vertical centering for now)
        page_offset = 0

        # Calculate how many cm to show before and after 0
        cm_before_zero = int(page_offset / pixels_per_cm) + 2
        cm_after_zero = int((self.height() - page_offset) / pixels_per_cm) + 2

        # Draw tick marks every mm, including negative numbers
        for i in range(-cm_before_zero * 10, cm_after_zero * 10):
            y = page_offset + int(i * pixels_per_cm / 10)
            if y < 0 or y > self.height():
                continue

            # Every cm: long tick + number
            if i % 10 == 0:
                painter.drawLine(self.width() - 10, y, self.width() - 1, y)
                # Draw cm number (rotated, can be negative)
                cm = i // 10
                painter.save()
                painter.translate(8, y + 10)
                painter.rotate(-90)
                painter.drawText(0, 0, str(cm))
                painter.restore()
            # Every 5mm: medium tick
            elif i % 5 == 0:
                painter.drawLine(self.width() - 7, y, self.width() - 1, y)
            # Every mm: small tick
            else:
                painter.drawLine(self.width() - 4, y, self.width() - 1, y)

    def set_zoom(self, zoom):
        """Set the zoom level."""
        self.zoom = zoom / 100.0
        self.update()
