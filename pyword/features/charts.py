"""Chart and Graph tools for PyWord - create and edit charts."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                               QTableWidget, QTableWidgetItem, QLabel, QGroupBox,
                               QComboBox, QSpinBox, QTabWidget, QWidget, QColorDialog,
                               QLineEdit, QCheckBox, QFormLayout, QHeaderView, QSplitter)
from PySide6.QtCore import Qt, Signal, QPointF, QRectF, QSize
from PySide6.QtGui import (QColor, QPainter, QPen, QBrush, QFont, QImage,
                          QTextImageFormat, QTextCursor, QPainterPath, QLinearGradient)
from enum import Enum
from typing import List, Tuple


class ChartType(Enum):
    """Available chart types."""
    BAR = "Bar Chart"
    LINE = "Line Chart"
    PIE = "Pie Chart"
    AREA = "Area Chart"
    SCATTER = "Scatter Plot"
    COLUMN = "Column Chart"
    DONUT = "Donut Chart"


class ChartData:
    """Holds chart data and configuration."""

    def __init__(self):
        self.title = "Chart Title"
        self.x_axis_label = "X Axis"
        self.y_axis_label = "Y Axis"
        self.categories = []
        self.series = []  # List of (name, values) tuples
        self.colors = []
        self.show_legend = True
        self.show_grid = True
        self.chart_type = ChartType.COLUMN

    def add_series(self, name: str, values: List[float], color: QColor = None):
        """Add a data series to the chart."""
        self.series.append((name, values))
        if color:
            self.colors.append(color)
        else:
            # Generate default color
            self.colors.append(self.generate_color(len(self.series) - 1))

    def generate_color(self, index: int) -> QColor:
        """Generate a color for a series based on index."""
        default_colors = [
            QColor(68, 114, 196),   # Blue
            QColor(237, 125, 49),   # Orange
            QColor(165, 165, 165),  # Gray
            QColor(255, 192, 0),    # Yellow
            QColor(91, 155, 213),   # Light Blue
            QColor(112, 173, 71),   # Green
            QColor(38, 68, 120),    # Dark Blue
            QColor(158, 72, 14),    # Dark Orange
        ]
        return default_colors[index % len(default_colors)]


class ChartRenderer:
    """Renders charts to images."""

    def __init__(self, width: int = 800, height: int = 600):
        self.width = width
        self.height = height
        self.margin = 50
        self.legend_width = 150

    def render_chart(self, data: ChartData) -> QImage:
        """Render a chart to a QImage."""
        image = QImage(self.width, self.height, QImage.Format_ARGB32)
        image.fill(Qt.white)

        painter = QPainter(image)
        painter.setRenderHint(QPainter.Antialiasing)

        # Render based on chart type
        if data.chart_type == ChartType.COLUMN:
            self.render_column_chart(painter, data)
        elif data.chart_type == ChartType.BAR:
            self.render_bar_chart(painter, data)
        elif data.chart_type == ChartType.LINE:
            self.render_line_chart(painter, data)
        elif data.chart_type == ChartType.PIE:
            self.render_pie_chart(painter, data)
        elif data.chart_type == ChartType.AREA:
            self.render_area_chart(painter, data)
        elif data.chart_type == ChartType.SCATTER:
            self.render_scatter_chart(painter, data)
        elif data.chart_type == ChartType.DONUT:
            self.render_donut_chart(painter, data)

        painter.end()
        return image

    def render_column_chart(self, painter: QPainter, data: ChartData):
        """Render a column chart."""
        # Calculate chart area
        chart_width = self.width - 2 * self.margin
        if data.show_legend:
            chart_width -= self.legend_width

        chart_height = self.height - 2 * self.margin - 50  # Extra space for title

        # Draw title
        painter.setFont(QFont("Arial", 14, QFont.Bold))
        painter.drawText(QRectF(0, 10, self.width, 30), Qt.AlignCenter, data.title)

        # Calculate max value for scaling
        max_value = 0
        for _, values in data.series:
            max_value = max(max_value, max(values) if values else 0)

        if max_value == 0:
            max_value = 1

        # Draw axes
        x_start = self.margin
        y_start = self.height - self.margin
        y_end = self.margin + 50

        painter.setPen(QPen(Qt.black, 2))
        painter.drawLine(x_start, y_start, x_start, y_end)  # Y-axis
        painter.drawLine(x_start, y_start, x_start + chart_width, y_start)  # X-axis

        # Draw grid
        if data.show_grid:
            painter.setPen(QPen(QColor(200, 200, 200), 1, Qt.DashLine))
            for i in range(1, 6):
                y = y_start - (chart_height * i / 5)
                painter.drawLine(x_start, y, x_start + chart_width, y)

        # Calculate bar dimensions
        num_categories = len(data.categories) if data.categories else 1
        num_series = len(data.series)

        category_width = chart_width / num_categories
        bar_width = category_width / (num_series + 1)

        # Draw bars
        for series_idx, (series_name, values) in enumerate(data.series):
            color = data.colors[series_idx] if series_idx < len(data.colors) else Qt.blue

            for cat_idx, value in enumerate(values):
                if cat_idx >= num_categories:
                    break

                # Calculate bar position and height
                bar_height = (value / max_value) * chart_height
                x = x_start + cat_idx * category_width + series_idx * bar_width + bar_width / 2
                y = y_start - bar_height

                # Draw bar with gradient
                gradient = QLinearGradient(x, y, x + bar_width, y)
                gradient.setColorAt(0, color)
                gradient.setColorAt(1, color.darker(110))

                painter.fillRect(QRectF(x, y, bar_width, bar_height), QBrush(gradient))
                painter.setPen(QPen(color.darker(120), 1))
                painter.drawRect(QRectF(x, y, bar_width, bar_height))

        # Draw category labels
        painter.setPen(Qt.black)
        painter.setFont(QFont("Arial", 9))
        for i, category in enumerate(data.categories):
            x = x_start + i * category_width + category_width / 2
            painter.drawText(QRectF(x - 50, y_start + 5, 100, 20),
                           Qt.AlignCenter, category)

        # Draw Y-axis labels
        for i in range(6):
            value = max_value * i / 5
            y = y_start - (chart_height * i / 5)
            painter.drawText(QRectF(5, y - 10, self.margin - 10, 20),
                           Qt.AlignRight | Qt.AlignVCenter, f"{value:.1f}")

        # Draw axis labels
        painter.setFont(QFont("Arial", 10, QFont.Bold))
        painter.drawText(QRectF(0, y_start + 30, self.width, 20),
                        Qt.AlignCenter, data.x_axis_label)

        painter.save()
        painter.translate(15, self.height / 2)
        painter.rotate(-90)
        painter.drawText(QRectF(-100, 0, 200, 20), Qt.AlignCenter, data.y_axis_label)
        painter.restore()

        # Draw legend
        if data.show_legend:
            self.draw_legend(painter, data, self.width - self.legend_width - 10, 60)

    def render_bar_chart(self, painter: QPainter, data: ChartData):
        """Render a horizontal bar chart."""
        # Similar to column chart but rotated
        # Implementation would be similar to column chart with x and y swapped
        painter.setFont(QFont("Arial", 12))
        painter.drawText(100, 100, "Bar Chart (Horizontal)")

    def render_line_chart(self, painter: QPainter, data: ChartData):
        """Render a line chart."""
        chart_width = self.width - 2 * self.margin
        if data.show_legend:
            chart_width -= self.legend_width

        chart_height = self.height - 2 * self.margin - 50

        # Draw title
        painter.setFont(QFont("Arial", 14, QFont.Bold))
        painter.drawText(QRectF(0, 10, self.width, 30), Qt.AlignCenter, data.title)

        # Calculate max value
        max_value = 0
        for _, values in data.series:
            max_value = max(max_value, max(values) if values else 0)

        if max_value == 0:
            max_value = 1

        x_start = self.margin
        y_start = self.height - self.margin
        y_end = self.margin + 50

        # Draw axes
        painter.setPen(QPen(Qt.black, 2))
        painter.drawLine(x_start, y_start, x_start, y_end)
        painter.drawLine(x_start, y_start, x_start + chart_width, y_start)

        # Draw lines
        num_points = len(data.categories) if data.categories else 1
        point_spacing = chart_width / (num_points - 1) if num_points > 1 else chart_width

        for series_idx, (series_name, values) in enumerate(data.series):
            color = data.colors[series_idx] if series_idx < len(data.colors) else Qt.blue

            painter.setPen(QPen(color, 3))

            # Draw line
            for i in range(len(values) - 1):
                x1 = x_start + i * point_spacing
                y1 = y_start - (values[i] / max_value) * chart_height

                x2 = x_start + (i + 1) * point_spacing
                y2 = y_start - (values[i + 1] / max_value) * chart_height

                painter.drawLine(QPointF(x1, y1), QPointF(x2, y2))

            # Draw points
            painter.setBrush(color)
            for i, value in enumerate(values):
                x = x_start + i * point_spacing
                y = y_start - (value / max_value) * chart_height
                painter.drawEllipse(QPointF(x, y), 4, 4)

        if data.show_legend:
            self.draw_legend(painter, data, self.width - self.legend_width - 10, 60)

    def render_pie_chart(self, painter: QPainter, data: ChartData):
        """Render a pie chart."""
        # Draw title
        painter.setFont(QFont("Arial", 14, QFont.Bold))
        painter.drawText(QRectF(0, 10, self.width, 30), Qt.AlignCenter, data.title)

        # Calculate total
        if not data.series or not data.series[0][1]:
            return

        values = data.series[0][1]  # Use first series for pie chart
        total = sum(values)

        if total == 0:
            return

        # Calculate pie dimensions
        pie_size = min(self.width, self.height) - 200
        pie_x = (self.width - pie_size) / 2
        pie_y = 60 + (self.height - 60 - pie_size) / 2

        # Draw pie slices
        start_angle = 0
        for i, value in enumerate(values):
            span_angle = int((value / total) * 360 * 16)  # Qt uses 1/16th degrees

            color = data.colors[i] if i < len(data.colors) else QColor(100, 100, 100)
            painter.setBrush(color)
            painter.setPen(QPen(Qt.white, 2))

            painter.drawPie(QRectF(pie_x, pie_y, pie_size, pie_size),
                          start_angle, span_angle)

            start_angle += span_angle

        # Draw legend
        if data.show_legend and data.categories:
            legend_x = self.width - self.legend_width - 10
            legend_y = 60
            self.draw_pie_legend(painter, data, values, total, legend_x, legend_y)

    def render_donut_chart(self, painter: QPainter, data: ChartData):
        """Render a donut chart (pie chart with hole in center)."""
        # Similar to pie chart but with inner circle
        self.render_pie_chart(painter, data)

        # Draw white circle in center
        pie_size = min(self.width, self.height) - 200
        pie_x = (self.width - pie_size) / 2
        pie_y = 60 + (self.height - 60 - pie_size) / 2

        inner_size = pie_size * 0.5
        inner_x = pie_x + (pie_size - inner_size) / 2
        inner_y = pie_y + (pie_size - inner_size) / 2

        painter.setBrush(Qt.white)
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(QRectF(inner_x, inner_y, inner_size, inner_size))

    def render_area_chart(self, painter: QPainter, data: ChartData):
        """Render an area chart."""
        # Similar to line chart but filled
        chart_width = self.width - 2 * self.margin
        if data.show_legend:
            chart_width -= self.legend_width

        chart_height = self.height - 2 * self.margin - 50

        painter.setFont(QFont("Arial", 14, QFont.Bold))
        painter.drawText(QRectF(0, 10, self.width, 30), Qt.AlignCenter, data.title)

        max_value = 0
        for _, values in data.series:
            max_value = max(max_value, max(values) if values else 0)

        if max_value == 0:
            max_value = 1

        x_start = self.margin
        y_start = self.height - self.margin

        num_points = len(data.categories) if data.categories else 1
        point_spacing = chart_width / (num_points - 1) if num_points > 1 else chart_width

        for series_idx, (series_name, values) in enumerate(data.series):
            color = data.colors[series_idx] if series_idx < len(data.colors) else Qt.blue

            # Create path for area
            path = QPainterPath()
            if values:
                x = x_start
                y = y_start - (values[0] / max_value) * chart_height
                path.moveTo(x, y_start)
                path.lineTo(x, y)

                for i, value in enumerate(values):
                    x = x_start + i * point_spacing
                    y = y_start - (value / max_value) * chart_height
                    path.lineTo(x, y)

                path.lineTo(x, y_start)
                path.closeSubpath()

            # Fill area
            area_color = QColor(color)
            area_color.setAlpha(128)
            painter.fillPath(path, area_color)

            # Draw line
            painter.setPen(QPen(color, 2))
            for i in range(len(values) - 1):
                x1 = x_start + i * point_spacing
                y1 = y_start - (values[i] / max_value) * chart_height

                x2 = x_start + (i + 1) * point_spacing
                y2 = y_start - (values[i + 1] / max_value) * chart_height

                painter.drawLine(QPointF(x1, y1), QPointF(x2, y2))

    def render_scatter_chart(self, painter: QPainter, data: ChartData):
        """Render a scatter plot."""
        painter.setFont(QFont("Arial", 12))
        painter.drawText(100, 100, "Scatter Plot")

    def draw_legend(self, painter: QPainter, data: ChartData, x: int, y: int):
        """Draw chart legend."""
        painter.setFont(QFont("Arial", 9))

        for i, (series_name, _) in enumerate(data.series):
            color = data.colors[i] if i < len(data.colors) else Qt.blue

            # Draw color box
            painter.fillRect(QRectF(x, y + i * 25, 15, 15), color)
            painter.setPen(Qt.black)
            painter.drawRect(QRectF(x, y + i * 25, 15, 15))

            # Draw series name
            painter.drawText(QRectF(x + 20, y + i * 25, 120, 15),
                           Qt.AlignVCenter, series_name)

    def draw_pie_legend(self, painter: QPainter, data: ChartData,
                       values: List[float], total: float, x: int, y: int):
        """Draw pie chart legend with percentages."""
        painter.setFont(QFont("Arial", 9))

        for i, (category, value) in enumerate(zip(data.categories, values)):
            color = data.colors[i] if i < len(data.colors) else Qt.gray

            percentage = (value / total) * 100 if total > 0 else 0

            # Draw color box
            painter.fillRect(QRectF(x, y + i * 25, 15, 15), color)
            painter.setPen(Qt.black)
            painter.drawRect(QRectF(x, y + i * 25, 15, 15))

            # Draw category and percentage
            text = f"{category} ({percentage:.1f}%)"
            painter.drawText(QRectF(x + 20, y + i * 25, 120, 15),
                           Qt.AlignVCenter, text)


class ChartEditor(QDialog):
    """Dialog for creating and editing charts."""

    chart_created = Signal(ChartData)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Chart")
        self.resize(900, 700)

        self.chart_data = ChartData()
        self.chart_data.categories = ["Category 1", "Category 2", "Category 3", "Category 4"]
        self.chart_data.add_series("Series 1", [4.3, 2.5, 3.5, 4.5])
        self.chart_data.add_series("Series 2", [2.4, 4.4, 1.8, 2.8])

        self.renderer = ChartRenderer()
        self.setup_ui()
        self.update_preview()

    def setup_ui(self):
        """Initialize the chart editor UI."""
        layout = QVBoxLayout(self)

        # Splitter for settings and preview
        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter)

        # Left side: Settings
        settings_widget = QWidget()
        settings_layout = QVBoxLayout(settings_widget)

        # Chart type selection
        type_group = QGroupBox("Chart Type")
        type_layout = QVBoxLayout(type_group)

        self.chart_type_combo = QComboBox()
        for chart_type in ChartType:
            self.chart_type_combo.addItem(chart_type.value, chart_type)
        self.chart_type_combo.currentIndexChanged.connect(self.on_chart_type_changed)
        type_layout.addWidget(self.chart_type_combo)

        settings_layout.addWidget(type_group)

        # Chart settings
        chart_group = QGroupBox("Chart Settings")
        chart_layout = QFormLayout(chart_group)

        self.title_input = QLineEdit(self.chart_data.title)
        self.title_input.textChanged.connect(self.on_settings_changed)
        chart_layout.addRow("Title:", self.title_input)

        self.x_axis_input = QLineEdit(self.chart_data.x_axis_label)
        self.x_axis_input.textChanged.connect(self.on_settings_changed)
        chart_layout.addRow("X-Axis Label:", self.x_axis_input)

        self.y_axis_input = QLineEdit(self.chart_data.y_axis_label)
        self.y_axis_input.textChanged.connect(self.on_settings_changed)
        chart_layout.addRow("Y-Axis Label:", self.y_axis_input)

        self.show_legend_check = QCheckBox()
        self.show_legend_check.setChecked(self.chart_data.show_legend)
        self.show_legend_check.stateChanged.connect(self.on_settings_changed)
        chart_layout.addRow("Show Legend:", self.show_legend_check)

        self.show_grid_check = QCheckBox()
        self.show_grid_check.setChecked(self.chart_data.show_grid)
        self.show_grid_check.stateChanged.connect(self.on_settings_changed)
        chart_layout.addRow("Show Grid:", self.show_grid_check)

        settings_layout.addWidget(chart_group)

        # Data table
        data_group = QGroupBox("Chart Data")
        data_layout = QVBoxLayout(data_group)

        self.data_table = QTableWidget(5, 3)
        self.data_table.setHorizontalHeaderLabels(["Category", "Series 1", "Series 2"])
        self.data_table.cellChanged.connect(self.on_data_changed)
        self.populate_data_table()
        data_layout.addWidget(self.data_table)

        # Data buttons
        data_button_layout = QHBoxLayout()
        add_row_button = QPushButton("Add Row")
        add_row_button.clicked.connect(self.add_row)
        data_button_layout.addWidget(add_row_button)

        add_series_button = QPushButton("Add Series")
        add_series_button.clicked.connect(self.add_series)
        data_button_layout.addWidget(add_series_button)

        data_layout.addLayout(data_button_layout)

        settings_layout.addWidget(data_group)
        settings_layout.addStretch()

        splitter.addWidget(settings_widget)

        # Right side: Preview
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)

        preview_label = QLabel("Preview:")
        preview_layout.addWidget(preview_label)

        self.preview_label = QLabel()
        self.preview_label.setMinimumSize(600, 450)
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

    def populate_data_table(self):
        """Populate the data table with current chart data."""
        self.data_table.blockSignals(True)

        for i, category in enumerate(self.chart_data.categories):
            self.data_table.setItem(i, 0, QTableWidgetItem(category))

            for series_idx, (_, values) in enumerate(self.chart_data.series):
                if i < len(values):
                    self.data_table.setItem(i, series_idx + 1,
                                          QTableWidgetItem(str(values[i])))

        self.data_table.blockSignals(False)

    def on_chart_type_changed(self):
        """Handle chart type change."""
        chart_type = self.chart_type_combo.currentData()
        self.chart_data.chart_type = chart_type
        self.update_preview()

    def on_settings_changed(self):
        """Handle settings change."""
        self.chart_data.title = self.title_input.text()
        self.chart_data.x_axis_label = self.x_axis_input.text()
        self.chart_data.y_axis_label = self.y_axis_input.text()
        self.chart_data.show_legend = self.show_legend_check.isChecked()
        self.chart_data.show_grid = self.show_grid_check.isChecked()
        self.update_preview()

    def on_data_changed(self, row: int, col: int):
        """Handle data table change."""
        self.update_chart_data_from_table()
        self.update_preview()

    def update_chart_data_from_table(self):
        """Update chart data from the data table."""
        # Update categories
        categories = []
        for row in range(self.data_table.rowCount()):
            item = self.data_table.item(row, 0)
            if item and item.text():
                categories.append(item.text())

        self.chart_data.categories = categories

        # Update series
        num_series = self.data_table.columnCount() - 1
        series_data = []

        for series_idx in range(num_series):
            values = []
            for row in range(len(categories)):
                item = self.data_table.item(row, series_idx + 1)
                try:
                    value = float(item.text()) if item and item.text() else 0.0
                except ValueError:
                    value = 0.0
                values.append(value)

            header = self.data_table.horizontalHeaderItem(series_idx + 1)
            series_name = header.text() if header else f"Series {series_idx + 1}"
            series_data.append((series_name, values))

        # Update series
        self.chart_data.series = series_data
        self.chart_data.colors = [self.chart_data.generate_color(i)
                                 for i in range(len(series_data))]

    def add_row(self):
        """Add a new row to the data table."""
        row_count = self.data_table.rowCount()
        self.data_table.insertRow(row_count)
        self.data_table.setItem(row_count, 0,
                               QTableWidgetItem(f"Category {row_count + 1}"))

    def add_series(self):
        """Add a new series to the data table."""
        col_count = self.data_table.columnCount()
        self.data_table.insertColumn(col_count)
        self.data_table.setHorizontalHeaderItem(col_count,
                                               QTableWidgetItem(f"Series {col_count}"))

    def update_preview(self):
        """Update the chart preview."""
        image = self.renderer.render_chart(self.chart_data)
        self.preview_label.setPixmap(image.scaled(
            self.preview_label.size(),
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        ))

    def get_chart_data(self) -> ChartData:
        """Get the current chart data."""
        return self.chart_data


class ChartManager:
    """Manages charts in a document."""

    def __init__(self, editor):
        self.editor = editor
        self.charts = {}
        self.renderer = ChartRenderer()

    def show_chart_editor(self):
        """Show the chart editor dialog."""
        dialog = ChartEditor(self.editor)
        if dialog.exec() == QDialog.Accepted:
            chart_data = dialog.get_chart_data()
            self.insert_chart(chart_data)

    def insert_chart(self, chart_data: ChartData):
        """Insert a chart into the document."""
        # Generate unique ID
        chart_id = f"chart_{len(self.charts)}"
        self.charts[chart_id] = chart_data

        # Render chart
        image = self.renderer.render_chart(chart_data)

        # Insert into document
        cursor = self.editor.textCursor()

        image_format = QTextImageFormat()
        image_format.setName(chart_id)

        self.editor.document().addResource(
            QTextCursor.ImageResource,
            chart_id,
            image
        )

        cursor.insertImage(image_format)
