"""Touch support for PyWord - enables touch gestures and interactions."""

from PySide6.QtWidgets import QWidget, QGestureRecognizer, QApplication
from PySide6.QtCore import Qt, QEvent, QPointF, Signal, QObject, QTimer
from PySide6.QtGui import (QTouchEvent, QGesture, QPinchGesture, QPanGesture,
                          QSwipeGesture, QTapGesture, QTapAndHoldGesture)


class TouchManager(QObject):
    """Manages touch interactions and gestures for the application."""

    # Signals for touch events
    pinch_zoom = Signal(float)  # Scale factor
    pan_gesture = Signal(QPointF)  # Delta movement
    swipe_gesture = Signal(str)  # Direction: left, right, up, down
    tap_gesture = Signal(QPointF)  # Position
    tap_and_hold_gesture = Signal(QPointF)  # Position
    touch_mode_changed = Signal(bool)  # Touch mode enabled/disabled

    def __init__(self, parent=None):
        super().__init__(parent)
        self.touch_mode_enabled = True
        self.touch_scroll_enabled = True
        self.gesture_enabled = True

        # Touch sensitivity settings
        self.pinch_sensitivity = 1.0
        self.pan_sensitivity = 1.0
        self.tap_hold_duration = 500  # milliseconds

    def enable_touch_mode(self, enabled: bool = True):
        """Enable or disable touch mode."""
        self.touch_mode_enabled = enabled
        self.touch_mode_changed.emit(enabled)

        if enabled:
            # Set application attribute for touch events
            QApplication.setAttribute(Qt.AA_SynthesizeTouchForUnhandledMouseEvents, True)
            QApplication.setAttribute(Qt.AA_SynthesizeMouseForUnhandledTouchEvents, True)
        else:
            QApplication.setAttribute(Qt.AA_SynthesizeTouchForUnhandledMouseEvents, False)
            QApplication.setAttribute(Qt.AA_SynthesizeMouseForUnhandledTouchEvents, False)

    def install_touch_handler(self, widget: QWidget):
        """Install touch event handler on a widget."""
        if not self.touch_mode_enabled:
            return

        # Accept touch events
        widget.setAttribute(Qt.WA_AcceptTouchEvents, True)

        # Enable gestures
        if self.gesture_enabled:
            widget.grabGesture(Qt.PinchGesture)
            widget.grabGesture(Qt.PanGesture)
            widget.grabGesture(Qt.SwipeGesture)
            widget.grabGesture(Qt.TapGesture)
            widget.grabGesture(Qt.TapAndHoldGesture)

        # Install event filter
        widget.installEventFilter(self)

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:
        """Filter and handle touch events."""
        if not self.touch_mode_enabled:
            return False

        # Handle touch events
        if event.type() == QEvent.TouchBegin:
            return self.handle_touch_begin(event)
        elif event.type() == QEvent.TouchUpdate:
            return self.handle_touch_update(event)
        elif event.type() == QEvent.TouchEnd:
            return self.handle_touch_end(event)

        # Handle gestures
        elif event.type() == QEvent.Gesture:
            return self.handle_gesture_event(event)

        return False

    def handle_touch_begin(self, event: QTouchEvent) -> bool:
        """Handle touch begin event."""
        # Accept the event
        event.accept()
        return True

    def handle_touch_update(self, event: QTouchEvent) -> bool:
        """Handle touch update event."""
        # Process touch points
        touch_points = event.points()

        # Single finger drag/scroll
        if len(touch_points) == 1:
            point = touch_points[0]
            if self.touch_scroll_enabled:
                # Calculate scroll delta
                last_pos = point.lastPosition()
                current_pos = point.position()
                delta = QPointF(
                    (current_pos.x() - last_pos.x()) * self.pan_sensitivity,
                    (current_pos.y() - last_pos.y()) * self.pan_sensitivity
                )
                self.pan_gesture.emit(delta)

        event.accept()
        return True

    def handle_touch_end(self, event: QTouchEvent) -> bool:
        """Handle touch end event."""
        event.accept()
        return True

    def handle_gesture_event(self, event) -> bool:
        """Handle gesture events."""
        gesture = event.gesture(Qt.PinchGesture)
        if gesture:
            self.handle_pinch_gesture(gesture)

        gesture = event.gesture(Qt.PanGesture)
        if gesture:
            self.handle_pan_gesture(gesture)

        gesture = event.gesture(Qt.SwipeGesture)
        if gesture:
            self.handle_swipe_gesture(gesture)

        gesture = event.gesture(Qt.TapGesture)
        if gesture:
            self.handle_tap_gesture(gesture)

        gesture = event.gesture(Qt.TapAndHoldGesture)
        if gesture:
            self.handle_tap_and_hold_gesture(gesture)

        event.accept()
        return True

    def handle_pinch_gesture(self, gesture: QPinchGesture):
        """Handle pinch gesture for zooming."""
        if gesture.state() == Qt.GestureState.GestureUpdated:
            scale_factor = gesture.scaleFactor() * self.pinch_sensitivity
            self.pinch_zoom.emit(scale_factor)

    def handle_pan_gesture(self, gesture: QPanGesture):
        """Handle pan gesture for scrolling."""
        if gesture.state() == Qt.GestureState.GestureUpdated:
            delta = gesture.delta() * self.pan_sensitivity
            self.pan_gesture.emit(delta)

    def handle_swipe_gesture(self, gesture: QSwipeGesture):
        """Handle swipe gesture."""
        if gesture.state() == Qt.GestureState.GestureFinished:
            horizontal_direction = gesture.horizontalDirection()
            vertical_direction = gesture.verticalDirection()

            # Determine primary swipe direction
            if horizontal_direction == QSwipeGesture.Left:
                self.swipe_gesture.emit("left")
            elif horizontal_direction == QSwipeGesture.Right:
                self.swipe_gesture.emit("right")
            elif vertical_direction == QSwipeGesture.Up:
                self.swipe_gesture.emit("up")
            elif vertical_direction == QSwipeGesture.Down:
                self.swipe_gesture.emit("down")

    def handle_tap_gesture(self, gesture: QTapGesture):
        """Handle tap gesture."""
        if gesture.state() == Qt.GestureState.GestureFinished:
            position = gesture.position()
            self.tap_gesture.emit(position)

    def handle_tap_and_hold_gesture(self, gesture: QTapAndHoldGesture):
        """Handle tap and hold gesture (context menu trigger)."""
        if gesture.state() == Qt.GestureState.GestureFinished:
            position = gesture.position()
            self.tap_and_hold_gesture.emit(position)

    def set_pinch_sensitivity(self, sensitivity: float):
        """Set pinch gesture sensitivity (1.0 = normal)."""
        self.pinch_sensitivity = max(0.1, min(5.0, sensitivity))

    def set_pan_sensitivity(self, sensitivity: float):
        """Set pan gesture sensitivity (1.0 = normal)."""
        self.pan_sensitivity = max(0.1, min(5.0, sensitivity))


class TouchOptimizedWidget(QWidget):
    """Base class for touch-optimized widgets."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.touch_manager = None
        self.setup_touch_support()

    def setup_touch_support(self):
        """Setup touch support for this widget."""
        # Accept touch events
        self.setAttribute(Qt.WA_AcceptTouchEvents, True)

        # Increase touch target size
        self.setMinimumSize(44, 44)  # Minimum touch target size

    def set_touch_manager(self, manager: TouchManager):
        """Set the touch manager for this widget."""
        self.touch_manager = manager
        if manager:
            manager.install_touch_handler(self)

    def event(self, event: QEvent) -> bool:
        """Handle events including touch events."""
        # Let touch manager handle touch events
        if self.touch_manager and event.type() in (
            QEvent.TouchBegin,
            QEvent.TouchUpdate,
            QEvent.TouchEnd,
            QEvent.Gesture
        ):
            return self.touch_manager.eventFilter(self, event)

        return super().event(event)


class TouchOptimizedScrollArea(TouchOptimizedWidget):
    """Touch-optimized scroll area with momentum scrolling."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.velocity = QPointF(0, 0)
        self.deceleration = 0.95
        self.momentum_timer = QTimer(self)
        self.momentum_timer.timeout.connect(self.apply_momentum)

    def handle_pan_end(self, velocity: QPointF):
        """Handle end of pan gesture with momentum."""
        self.velocity = velocity
        if not self.momentum_timer.isActive():
            self.momentum_timer.start(16)  # ~60 FPS

    def apply_momentum(self):
        """Apply momentum scrolling."""
        # Apply velocity
        if abs(self.velocity.x()) > 0.1 or abs(self.velocity.y()) > 0.1:
            # Scroll by velocity
            # (Implementation would scroll the actual scroll area)

            # Apply deceleration
            self.velocity *= self.deceleration
        else:
            # Stop momentum
            self.velocity = QPointF(0, 0)
            self.momentum_timer.stop()


def make_touch_friendly(widget: QWidget, min_size: int = 44):
    """Make a widget more touch-friendly by increasing its size."""
    current_size = widget.minimumSize()
    new_width = max(current_size.width(), min_size)
    new_height = max(current_size.height(), min_size)
    widget.setMinimumSize(new_width, new_height)

    # Increase padding for better touch targets
    current_style = widget.styleSheet()
    touch_style = f"{current_style}\npadding: 8px;"
    widget.setStyleSheet(touch_style)


def enable_kinetic_scrolling(scroll_area):
    """Enable kinetic (momentum) scrolling for a scroll area."""
    # This would be implemented with custom scroll behavior
    # For now, just enable touch events
    scroll_area.setAttribute(Qt.WA_AcceptTouchEvents, True)
    scroll_area.grabGesture(Qt.PanGesture)
