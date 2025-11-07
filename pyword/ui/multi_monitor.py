"""Multi-monitor support for PyWord - manage windows across multiple displays."""

from PySide6.QtWidgets import QApplication, QWidget, QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QComboBox, QCheckBox, QGroupBox
from PySide6.QtCore import Qt, Signal, QObject, QSettings, QRect, QPoint
from PySide6.QtGui import QScreen, QWindow
from typing import Optional, List


class MonitorInfo:
    """Information about a monitor/screen."""

    def __init__(self, screen: QScreen, index: int):
        self.screen = screen
        self.index = index
        self.name = screen.name()
        self.geometry = screen.geometry()
        self.available_geometry = screen.availableGeometry()
        self.size = screen.size()
        self.physical_size = screen.physicalSize()
        self.dpi = screen.logicalDotsPerInch()
        self.refresh_rate = screen.refreshRate()
        self.is_primary = (screen == QApplication.primaryScreen())

    def __str__(self) -> str:
        """String representation of monitor info."""
        primary = " (Primary)" if self.is_primary else ""
        return f"{self.name}{primary} - {self.size.width()}x{self.size.height()} @ {self.dpi} DPI"


class MultiMonitorManager(QObject):
    """Manages multi-monitor setup and window positioning."""

    monitor_added = Signal(MonitorInfo)
    monitor_removed = Signal(MonitorInfo)
    primary_monitor_changed = Signal(MonitorInfo)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings = QSettings("PyWord", "MultiMonitor")
        self.monitors: List[MonitorInfo] = []
        self.window_positions = {}  # Store window positions for each monitor

        # Connect to screen changes
        app = QApplication.instance()
        if app:
            app.screenAdded.connect(self.on_screen_added)
            app.screenRemoved.connect(self.on_screen_removed)
            app.primaryScreenChanged.connect(self.on_primary_screen_changed)

        # Initialize monitors
        self.refresh_monitors()

    def refresh_monitors(self):
        """Refresh the list of available monitors."""
        self.monitors.clear()
        screens = QApplication.screens()

        for index, screen in enumerate(screens):
            monitor_info = MonitorInfo(screen, index)
            self.monitors.append(monitor_info)

    def get_monitors(self) -> List[MonitorInfo]:
        """Get list of all available monitors."""
        return self.monitors

    def get_monitor_count(self) -> int:
        """Get the number of available monitors."""
        return len(self.monitors)

    def get_primary_monitor(self) -> Optional[MonitorInfo]:
        """Get the primary monitor."""
        for monitor in self.monitors:
            if monitor.is_primary:
                return monitor
        return self.monitors[0] if self.monitors else None

    def get_monitor_at_point(self, point: QPoint) -> Optional[MonitorInfo]:
        """Get the monitor containing the specified point."""
        for monitor in self.monitors:
            if monitor.geometry.contains(point):
                return monitor
        return self.get_primary_monitor()

    def get_monitor_by_index(self, index: int) -> Optional[MonitorInfo]:
        """Get monitor by index."""
        if 0 <= index < len(self.monitors):
            return self.monitors[index]
        return None

    def move_window_to_monitor(self, window: QWidget, monitor_index: int,
                               maximize: bool = False):
        """Move a window to the specified monitor."""
        monitor = self.get_monitor_by_index(monitor_index)
        if not monitor:
            return

        # Get the available geometry (excludes taskbar, etc.)
        geometry = monitor.available_geometry

        if maximize:
            # Maximize on the target monitor
            window.setGeometry(geometry)
            window.showMaximized()
        else:
            # Center on the target monitor
            window_size = window.size()
            x = geometry.x() + (geometry.width() - window_size.width()) // 2
            y = geometry.y() + (geometry.height() - window_size.height()) // 2
            window.move(x, y)

    def center_window_on_monitor(self, window: QWidget, monitor_index: int = None):
        """Center a window on the specified monitor (or current monitor)."""
        if monitor_index is None:
            # Use current monitor
            monitor = self.get_monitor_at_point(window.pos())
        else:
            monitor = self.get_monitor_by_index(monitor_index)

        if not monitor:
            return

        geometry = monitor.available_geometry
        window_size = window.size()

        x = geometry.x() + (geometry.width() - window_size.width()) // 2
        y = geometry.y() + (geometry.height() - window_size.height()) // 2

        window.move(x, y)

    def save_window_position(self, window_id: str, window: QWidget):
        """Save window position and geometry."""
        geometry = window.geometry()
        monitor = self.get_monitor_at_point(window.pos())

        position_data = {
            'x': geometry.x(),
            'y': geometry.y(),
            'width': geometry.width(),
            'height': geometry.height(),
            'monitor_index': monitor.index if monitor else 0,
            'maximized': window.isMaximized()
        }

        self.settings.setValue(f"window_{window_id}", position_data)

    def restore_window_position(self, window_id: str, window: QWidget,
                                default_monitor: int = 0):
        """Restore window position and geometry."""
        position_data = self.settings.value(f"window_{window_id}")

        if position_data:
            # Check if the saved monitor still exists
            monitor_index = position_data.get('monitor_index', default_monitor)
            monitor = self.get_monitor_by_index(monitor_index)

            if monitor:
                # Restore position
                geometry = QRect(
                    position_data['x'],
                    position_data['y'],
                    position_data['width'],
                    position_data['height']
                )

                # Ensure window is visible on the monitor
                if monitor.available_geometry.intersects(geometry):
                    window.setGeometry(geometry)

                    if position_data.get('maximized', False):
                        window.showMaximized()
                else:
                    # Position is off-screen, center on monitor
                    self.center_window_on_monitor(window, monitor_index)
            else:
                # Monitor no longer exists, center on primary
                self.center_window_on_monitor(window, 0)
        else:
            # No saved position, center on default monitor
            self.center_window_on_monitor(window, default_monitor)

    def span_window_across_monitors(self, window: QWidget,
                                     monitor_indices: List[int]):
        """Span a window across multiple monitors."""
        if not monitor_indices:
            return

        # Calculate combined geometry
        min_x = min_y = float('inf')
        max_x = max_y = float('-inf')

        for index in monitor_indices:
            monitor = self.get_monitor_by_index(index)
            if monitor:
                geometry = monitor.available_geometry
                min_x = min(min_x, geometry.x())
                min_y = min(min_y, geometry.y())
                max_x = max(max_x, geometry.x() + geometry.width())
                max_y = max(max_y, geometry.y() + geometry.height())

        # Set window to span the calculated area
        window.setGeometry(int(min_x), int(min_y),
                          int(max_x - min_x), int(max_y - min_y))

    def on_screen_added(self, screen: QScreen):
        """Handle screen added event."""
        self.refresh_monitors()
        monitor = MonitorInfo(screen, len(self.monitors) - 1)
        self.monitor_added.emit(monitor)

    def on_screen_removed(self, screen: QScreen):
        """Handle screen removed event."""
        # Find and emit the removed monitor before refreshing
        for monitor in self.monitors:
            if monitor.screen == screen:
                self.monitor_removed.emit(monitor)
                break

        self.refresh_monitors()

    def on_primary_screen_changed(self, screen: QScreen):
        """Handle primary screen changed event."""
        self.refresh_monitors()
        for monitor in self.monitors:
            if monitor.screen == screen:
                self.primary_monitor_changed.emit(monitor)
                break

    def get_monitor_info_string(self, monitor: MonitorInfo) -> str:
        """Get detailed information string for a monitor."""
        info = [
            f"Monitor: {monitor.name}",
            f"Resolution: {monitor.size.width()}x{monitor.size.height()}",
            f"Position: ({monitor.geometry.x()}, {monitor.geometry.y()})",
            f"DPI: {monitor.dpi}",
            f"Refresh Rate: {monitor.refresh_rate} Hz",
            f"Physical Size: {monitor.physical_size.width():.1f}mm x {monitor.physical_size.height():.1f}mm"
        ]

        if monitor.is_primary:
            info.append("Primary Monitor")

        return "\n".join(info)


class MonitorSelectionDialog(QDialog):
    """Dialog for selecting a monitor for window placement."""

    def __init__(self, monitor_manager: MultiMonitorManager, parent=None):
        super().__init__(parent)
        self.monitor_manager = monitor_manager
        self.selected_monitor_index = 0
        self.setup_ui()

    def setup_ui(self):
        """Initialize the dialog UI."""
        self.setWindowTitle("Select Monitor")
        self.setModal(True)
        layout = QVBoxLayout(self)

        # Monitor selection
        monitor_group = QGroupBox("Available Monitors")
        monitor_layout = QVBoxLayout(monitor_group)

        self.monitor_combo = QComboBox()
        for monitor in self.monitor_manager.get_monitors():
            self.monitor_combo.addItem(str(monitor), monitor.index)

        monitor_layout.addWidget(self.monitor_combo)

        # Monitor info label
        self.info_label = QLabel()
        self.update_monitor_info()
        self.monitor_combo.currentIndexChanged.connect(self.update_monitor_info)
        monitor_layout.addWidget(self.info_label)

        layout.addWidget(monitor_group)

        # Options
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout(options_group)

        self.maximize_checkbox = QCheckBox("Maximize window on monitor")
        options_layout.addWidget(self.maximize_checkbox)

        layout.addWidget(options_group)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        button_layout.addWidget(ok_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

    def update_monitor_info(self):
        """Update the monitor information display."""
        index = self.monitor_combo.currentData()
        if index is not None:
            monitor = self.monitor_manager.get_monitor_by_index(index)
            if monitor:
                info = self.monitor_manager.get_monitor_info_string(monitor)
                self.info_label.setText(info)

    def get_selected_monitor_index(self) -> int:
        """Get the selected monitor index."""
        return self.monitor_combo.currentData()

    def should_maximize(self) -> bool:
        """Check if window should be maximized."""
        return self.maximize_checkbox.isChecked()
