"""
Performance Settings Dialog

This dialog allows users to configure performance features.
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QGroupBox,
                              QCheckBox, QLabel, QPushButton, QSpinBox,
                              QTabWidget, QWidget, QTextEdit, QProgressBar)
from PySide6.QtCore import Qt


class PerformanceDialog(QDialog):
    """Dialog for configuring performance settings and viewing metrics."""

    def __init__(self, performance_manager, memory_manager, auto_recovery, parent=None):
        super().__init__(parent)
        self.performance_manager = performance_manager
        self.memory_manager = memory_manager
        self.auto_recovery = auto_recovery
        self.setWindowTitle("Performance Settings")
        self.setMinimumSize(600, 500)
        self.setup_ui()
        self.load_current_settings()

    def setup_ui(self):
        """Setup the dialog UI."""
        layout = QVBoxLayout(self)

        # Create tab widget
        tab_widget = QTabWidget()

        # Performance tab
        performance_tab = self.create_performance_tab()
        tab_widget.addTab(performance_tab, "Performance")

        # Memory tab
        memory_tab = self.create_memory_tab()
        tab_widget.addTab(memory_tab, "Memory")

        # Auto-Recovery tab
        recovery_tab = self.create_recovery_tab()
        tab_widget.addTab(recovery_tab, "Auto-Recovery")

        # Metrics tab
        metrics_tab = self.create_metrics_tab()
        tab_widget.addTab(metrics_tab, "Metrics")

        layout.addWidget(tab_widget)

        # Buttons
        button_layout = QHBoxLayout()

        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)

        self.apply_button = QPushButton("Apply")
        self.apply_button.clicked.connect(self.apply_settings)

        button_layout.addStretch()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.apply_button)

        layout.addLayout(button_layout)

    def create_performance_tab(self):
        """Create the performance settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Large document optimization
        group = QGroupBox("Large Document Optimization")
        group_layout = QVBoxLayout()

        self.optimize_large_docs = QCheckBox("Enable large document optimization")
        self.optimize_large_docs.setChecked(True)
        group_layout.addWidget(self.optimize_large_docs)

        chunk_layout = QHBoxLayout()
        chunk_layout.addWidget(QLabel("Chunk size (characters):"))

        self.chunk_size_spin = QSpinBox()
        self.chunk_size_spin.setRange(1000, 100000)
        self.chunk_size_spin.setValue(10000)
        self.chunk_size_spin.setSingleStep(1000)
        chunk_layout.addWidget(self.chunk_size_spin)

        group_layout.addLayout(chunk_layout)
        group.setLayout(group_layout)
        layout.addWidget(group)

        # Background saving
        group2 = QGroupBox("Background Saving")
        group2_layout = QVBoxLayout()

        self.enable_background_save = QCheckBox("Enable background saving")
        self.enable_background_save.setChecked(True)
        group2_layout.addWidget(self.enable_background_save)

        group2.setLayout(group2_layout)
        layout.addWidget(group2)

        # Rendering optimizations
        group3 = QGroupBox("Rendering")
        group3_layout = QVBoxLayout()

        self.lazy_rendering = QCheckBox("Enable lazy rendering for long documents")
        self.lazy_rendering.setChecked(True)
        group3_layout.addWidget(self.lazy_rendering)

        self.cache_rendered_pages = QCheckBox("Cache rendered pages")
        self.cache_rendered_pages.setChecked(True)
        group3_layout.addWidget(self.cache_rendered_pages)

        group3.setLayout(group3_layout)
        layout.addWidget(group3)

        layout.addStretch()
        return widget

    def create_memory_tab(self):
        """Create the memory management tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Memory limit
        limit_layout = QHBoxLayout()
        limit_layout.addWidget(QLabel("Maximum memory usage:"))

        self.memory_limit_spin = QSpinBox()
        self.memory_limit_spin.setRange(100, 4096)
        self.memory_limit_spin.setValue(500)
        self.memory_limit_spin.setSuffix(" MB")
        limit_layout.addWidget(self.memory_limit_spin)

        layout.addLayout(limit_layout)

        # Memory status
        group = QGroupBox("Current Memory Usage")
        group_layout = QVBoxLayout()

        self.memory_usage_label = QLabel("Calculating...")
        group_layout.addWidget(self.memory_usage_label)

        self.memory_progress = QProgressBar()
        self.memory_progress.setRange(0, 100)
        group_layout.addWidget(self.memory_progress)

        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self.update_memory_usage)
        group_layout.addWidget(refresh_btn)

        group.setLayout(group_layout)
        layout.addWidget(group)

        # Cache management
        group2 = QGroupBox("Cache Management")
        group2_layout = QVBoxLayout()

        self.auto_clear_cache = QCheckBox("Automatically clear cache when memory is low")
        self.auto_clear_cache.setChecked(True)
        group2_layout.addWidget(self.auto_clear_cache)

        clear_cache_btn = QPushButton("Clear Cache Now")
        clear_cache_btn.clicked.connect(self.clear_cache)
        group2_layout.addWidget(clear_cache_btn)

        group2.setLayout(group2_layout)
        layout.addWidget(group2)

        layout.addStretch()
        return widget

    def create_recovery_tab(self):
        """Create the auto-recovery settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Enable auto-recovery
        self.enable_auto_recovery = QCheckBox("Enable auto-recovery")
        self.enable_auto_recovery.setChecked(True)
        layout.addWidget(self.enable_auto_recovery)

        # Backup interval
        interval_layout = QHBoxLayout()
        interval_layout.addWidget(QLabel("Backup interval:"))

        self.backup_interval_spin = QSpinBox()
        self.backup_interval_spin.setRange(60, 3600)
        self.backup_interval_spin.setValue(300)
        self.backup_interval_spin.setSuffix(" seconds")
        interval_layout.addWidget(self.backup_interval_spin)

        layout.addLayout(interval_layout)

        # Backup location
        location_layout = QHBoxLayout()
        location_layout.addWidget(QLabel("Backup location:"))

        self.backup_location_label = QLabel(self.auto_recovery.backup_dir)
        self.backup_location_label.setWordWrap(True)
        location_layout.addWidget(self.backup_location_label)

        layout.addLayout(location_layout)

        # Recovery options
        group = QGroupBox("Recovery Options")
        group_layout = QVBoxLayout()

        self.auto_check_recovery = QCheckBox("Automatically check for recoverable files on startup")
        self.auto_check_recovery.setChecked(True)
        group_layout.addWidget(self.auto_check_recovery)

        self.keep_backup_after_save = QCheckBox("Keep backup files after saving")
        self.keep_backup_after_save.setChecked(False)
        group_layout.addWidget(self.keep_backup_after_save)

        clear_backups_btn = QPushButton("Clear All Backup Files")
        clear_backups_btn.clicked.connect(self.clear_all_backups)
        group_layout.addWidget(clear_backups_btn)

        group.setLayout(group_layout)
        layout.addWidget(group)

        layout.addStretch()
        return widget

    def create_metrics_tab(self):
        """Create the performance metrics tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        layout.addWidget(QLabel("Performance Metrics:"))

        self.metrics_text = QTextEdit()
        self.metrics_text.setReadOnly(True)
        layout.addWidget(self.metrics_text)

        refresh_btn = QPushButton("Refresh Metrics")
        refresh_btn.clicked.connect(self.update_metrics)
        layout.addWidget(refresh_btn)

        return widget

    def load_current_settings(self):
        """Load current settings."""
        # Update memory usage
        self.update_memory_usage()

        # Update metrics
        self.update_metrics()

    def apply_settings(self):
        """Apply the current settings."""
        # Apply memory limit
        max_memory = self.memory_limit_spin.value()
        self.memory_manager.max_memory_bytes = max_memory * 1024 * 1024

        # Apply auto-recovery settings
        if self.enable_auto_recovery.isChecked():
            interval = self.backup_interval_spin.value()
            self.auto_recovery.backup_interval = interval
            self.auto_recovery.start()
        else:
            self.auto_recovery.stop()

    def accept(self):
        """Apply settings and close dialog."""
        self.apply_settings()
        super().accept()

    def update_memory_usage(self):
        """Update memory usage display."""
        try:
            current_usage = self.memory_manager.get_memory_usage()
            max_usage = self.memory_manager.max_memory_bytes

            usage_mb = current_usage / (1024 * 1024)
            max_mb = max_usage / (1024 * 1024)

            self.memory_usage_label.setText(
                f"Current: {usage_mb:.1f} MB / {max_mb:.1f} MB"
            )

            percentage = int((current_usage / max_usage) * 100)
            self.memory_progress.setValue(percentage)

        except Exception as e:
            self.memory_usage_label.setText(f"Error: {str(e)}")

    def clear_cache(self):
        """Clear the memory cache."""
        self.memory_manager.clear_cache()
        self.update_memory_usage()

        from PySide6.QtWidgets import QMessageBox
        QMessageBox.information(self, "Cache Cleared", "Memory cache has been cleared.")

    def clear_all_backups(self):
        """Clear all backup files."""
        from PySide6.QtWidgets import QMessageBox
        reply = QMessageBox.question(
            self,
            "Clear Backups",
            "Are you sure you want to delete all backup files?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.auto_recovery.clear_all_backups()
            QMessageBox.information(self, "Backups Cleared", "All backup files have been deleted.")

    def update_metrics(self):
        """Update performance metrics display."""
        metrics = self.performance_manager.get_metrics_summary()

        text = "Performance Metrics Summary\n"
        text += "=" * 40 + "\n\n"

        if not metrics:
            text += "No metrics available yet.\n"
        else:
            for metric_name, stats in metrics.items():
                text += f"{metric_name.replace('_', ' ').title()}:\n"
                text += f"  Average: {stats['average']:.3f}\n"
                text += f"  Min: {stats['min']:.3f}\n"
                text += f"  Max: {stats['max']:.3f}\n"
                text += f"  Samples: {stats['count']}\n\n"

        self.metrics_text.setPlainText(text)
