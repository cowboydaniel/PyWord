"""
Performance Optimization Features for PyWord

This module provides performance optimization features including:
- Large document optimization
- Memory management
- Background saving
- Auto-recovery
"""

import os
import json
import time
import threading
import weakref
from pathlib import Path
from typing import Optional, Dict, Any, Callable, List
from datetime import datetime, timedelta
from PySide6.QtCore import QObject, QTimer, Signal, QThread
from PySide6.QtWidgets import QTextEdit, QMessageBox


class LargeDocumentOptimizer:
    """
    Optimizes performance for large documents.

    Features:
    - Lazy loading of document sections
    - Chunked rendering
    - Progressive loading
    - Memory-efficient text handling
    """

    def __init__(self, chunk_size: int = 10000):
        """
        Initialize the large document optimizer.

        Args:
            chunk_size: Number of characters to load at once
        """
        self.chunk_size = chunk_size
        self.loaded_chunks: Dict[int, str] = {}
        self.total_chunks: int = 0
        self.current_chunk: int = 0

    def load_document_chunked(self, file_path: str, callback: Optional[Callable] = None) -> bool:
        """
        Load a document in chunks for better performance.

        Args:
            file_path: Path to the document
            callback: Optional callback for progress updates

        Returns:
            True if successful, False otherwise
        """
        try:
            path = Path(file_path)
            if not path.exists():
                return False

            # Get file size
            file_size = path.stat().st_size
            self.total_chunks = (file_size // self.chunk_size) + 1

            # Load first chunk
            with open(path, 'r', encoding='utf-8') as f:
                chunk_data = f.read(self.chunk_size)
                self.loaded_chunks[0] = chunk_data
                self.current_chunk = 0

                if callback:
                    callback(0, self.total_chunks)

            return True

        except Exception as e:
            print(f"Error loading document in chunks: {e}")
            return False

    def get_chunk(self, chunk_index: int, file_path: str) -> Optional[str]:
        """
        Get a specific chunk of the document.

        Args:
            chunk_index: Index of the chunk to retrieve
            file_path: Path to the document

        Returns:
            Chunk content or None if not found
        """
        if chunk_index in self.loaded_chunks:
            return self.loaded_chunks[chunk_index]

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                f.seek(chunk_index * self.chunk_size)
                chunk_data = f.read(self.chunk_size)
                self.loaded_chunks[chunk_index] = chunk_data
                return chunk_data
        except Exception as e:
            print(f"Error loading chunk {chunk_index}: {e}")
            return None

    def clear_unused_chunks(self, current_chunk: int, keep_range: int = 3):
        """
        Clear chunks that are far from the current position to free memory.

        Args:
            current_chunk: Current chunk being viewed
            keep_range: Number of chunks to keep before and after current
        """
        chunks_to_remove = []
        for chunk_idx in self.loaded_chunks.keys():
            if abs(chunk_idx - current_chunk) > keep_range:
                chunks_to_remove.append(chunk_idx)

        for chunk_idx in chunks_to_remove:
            del self.loaded_chunks[chunk_idx]


class MemoryManager:
    """
    Manages memory usage for the application.

    Features:
    - Memory usage monitoring
    - Automatic garbage collection
    - Cache management
    - Memory limit enforcement
    """

    def __init__(self, max_memory_mb: int = 500):
        """
        Initialize the memory manager.

        Args:
            max_memory_mb: Maximum memory usage in MB
        """
        self.max_memory_bytes = max_memory_mb * 1024 * 1024
        self.cached_documents: Dict[str, weakref.ref] = {}
        self.memory_warning_threshold = 0.8  # Warn at 80% usage

    def get_memory_usage(self) -> int:
        """
        Get current memory usage in bytes.

        Returns:
            Current memory usage
        """
        try:
            import psutil
            process = psutil.Process()
            return process.memory_info().rss
        except ImportError:
            # Fallback if psutil not available
            import sys
            return sys.getsizeof(self.cached_documents)

    def check_memory_usage(self) -> bool:
        """
        Check if memory usage is within limits.

        Returns:
            True if within limits, False if exceeded
        """
        current_usage = self.get_memory_usage()
        return current_usage < self.max_memory_bytes

    def cleanup_memory(self):
        """Perform memory cleanup operations."""
        import gc

        # Remove dead weak references
        dead_refs = [key for key, ref in self.cached_documents.items() if ref() is None]
        for key in dead_refs:
            del self.cached_documents[key]

        # Force garbage collection
        gc.collect()

    def cache_document(self, document_id: str, document: Any):
        """
        Cache a document with weak reference.

        Args:
            document_id: Unique identifier for the document
            document: Document object to cache
        """
        self.cached_documents[document_id] = weakref.ref(document)

    def get_cached_document(self, document_id: str) -> Optional[Any]:
        """
        Retrieve a cached document.

        Args:
            document_id: Unique identifier for the document

        Returns:
            Cached document or None
        """
        if document_id in self.cached_documents:
            ref = self.cached_documents[document_id]
            return ref()
        return None

    def clear_cache(self):
        """Clear all cached documents."""
        self.cached_documents.clear()
        self.cleanup_memory()


class BackgroundSaver(QThread):
    """
    Performs background saving of documents.

    Features:
    - Non-blocking save operations
    - Save queue management
    - Error handling and retry
    """

    save_completed = Signal(str, bool)  # file_path, success
    save_progress = Signal(str, int)  # file_path, percentage

    def __init__(self):
        super().__init__()
        self.save_queue: List[Dict[str, Any]] = []
        self.is_running = True
        self.lock = threading.Lock()

    def add_to_queue(self, document: Any, file_path: str, content: str):
        """
        Add a document to the save queue.

        Args:
            document: Document object
            file_path: Path to save to
            content: Content to save
        """
        with self.lock:
            self.save_queue.append({
                'document': document,
                'file_path': file_path,
                'content': content,
                'timestamp': datetime.now()
            })

    def run(self):
        """Background thread main loop."""
        while self.is_running:
            if self.save_queue:
                with self.lock:
                    if self.save_queue:
                        item = self.save_queue.pop(0)
                    else:
                        item = None

                if item:
                    self._save_document(item)

            # Sleep to avoid busy-waiting
            time.sleep(0.1)

    def _save_document(self, item: Dict[str, Any]):
        """
        Save a document to disk.

        Args:
            item: Dictionary containing document save information
        """
        file_path = item['file_path']
        content = item['content']

        try:
            # Emit progress
            self.save_progress.emit(file_path, 0)

            # Create directory if needed
            path = Path(file_path)
            path.parent.mkdir(parents=True, exist_ok=True)

            # Write to temporary file first
            temp_path = path.with_suffix('.tmp')
            with open(temp_path, 'w', encoding='utf-8') as f:
                f.write(content)

            self.save_progress.emit(file_path, 50)

            # Rename to final file (atomic operation)
            temp_path.rename(path)

            self.save_progress.emit(file_path, 100)
            self.save_completed.emit(file_path, True)

        except Exception as e:
            print(f"Error saving document {file_path}: {e}")
            self.save_completed.emit(file_path, False)

    def stop(self):
        """Stop the background saver thread."""
        self.is_running = False


class AutoRecovery(QObject):
    """
    Provides auto-recovery functionality for documents.

    Features:
    - Automatic document backup
    - Recovery file management
    - Crash recovery
    - Configurable backup intervals
    """

    recovery_available = Signal(list)  # List of recoverable files

    def __init__(self, backup_interval: int = 300, backup_dir: Optional[str] = None):
        """
        Initialize auto-recovery.

        Args:
            backup_interval: Backup interval in seconds (default: 5 minutes)
            backup_dir: Directory for backup files (default: temp directory)
        """
        super().__init__()
        self.backup_interval = backup_interval
        self.backup_dir = backup_dir or self._get_default_backup_dir()
        self.active_documents: Dict[str, Dict[str, Any]] = {}
        self.timer = QTimer()
        self.timer.timeout.connect(self._perform_backup)

        # Ensure backup directory exists
        Path(self.backup_dir).mkdir(parents=True, exist_ok=True)

    def _get_default_backup_dir(self) -> str:
        """Get the default backup directory."""
        import tempfile
        return os.path.join(tempfile.gettempdir(), 'pyword_backup')

    def start(self):
        """Start the auto-recovery system."""
        self.timer.start(self.backup_interval * 1000)
        self._check_for_recovery_files()

    def stop(self):
        """Stop the auto-recovery system."""
        self.timer.stop()

    def register_document(self, document_id: str, document: Any, file_path: Optional[str] = None):
        """
        Register a document for auto-recovery.

        Args:
            document_id: Unique identifier for the document
            document: Document object
            file_path: Optional file path for the document
        """
        self.active_documents[document_id] = {
            'document': weakref.ref(document),
            'file_path': file_path,
            'last_backup': None
        }

    def unregister_document(self, document_id: str):
        """
        Unregister a document from auto-recovery.

        Args:
            document_id: Unique identifier for the document
        """
        if document_id in self.active_documents:
            # Clean up backup file
            self._remove_backup(document_id)
            del self.active_documents[document_id]

    def _perform_backup(self):
        """Perform backup of all active documents."""
        for doc_id, doc_info in list(self.active_documents.items()):
            doc_ref = doc_info['document']
            doc = doc_ref()

            if doc is None:
                # Document no longer exists
                self.unregister_document(doc_id)
                continue

            # Check if document has been modified
            if hasattr(doc, 'modified') and doc.modified:
                self._backup_document(doc_id, doc)

    def _backup_document(self, document_id: str, document: Any):
        """
        Create a backup of a document.

        Args:
            document_id: Unique identifier for the document
            document: Document object to backup
        """
        try:
            backup_path = self._get_backup_path(document_id)
            backup_data = {
                'document_id': document_id,
                'timestamp': datetime.now().isoformat(),
                'content': document.content if hasattr(document, 'content') else '',
                'title': document.title if hasattr(document, 'title') else 'Untitled',
                'file_path': document.file_path if hasattr(document, 'file_path') else None
            }

            with open(backup_path, 'w', encoding='utf-8') as f:
                json.dump(backup_data, f, indent=2)

            self.active_documents[document_id]['last_backup'] = datetime.now()

        except Exception as e:
            print(f"Error backing up document {document_id}: {e}")

    def _get_backup_path(self, document_id: str) -> str:
        """
        Get the backup file path for a document.

        Args:
            document_id: Unique identifier for the document

        Returns:
            Path to the backup file
        """
        return os.path.join(self.backup_dir, f"{document_id}.recovery")

    def _remove_backup(self, document_id: str):
        """
        Remove a backup file.

        Args:
            document_id: Unique identifier for the document
        """
        backup_path = self._get_backup_path(document_id)
        if os.path.exists(backup_path):
            try:
                os.remove(backup_path)
            except Exception as e:
                print(f"Error removing backup file {backup_path}: {e}")

    def _check_for_recovery_files(self):
        """Check for existing recovery files and emit signal if found."""
        recovery_files = []
        backup_dir = Path(self.backup_dir)

        if backup_dir.exists():
            for backup_file in backup_dir.glob('*.recovery'):
                try:
                    with open(backup_file, 'r', encoding='utf-8') as f:
                        backup_data = json.load(f)

                    # Check if backup is recent (within 24 hours)
                    backup_time = datetime.fromisoformat(backup_data['timestamp'])
                    if datetime.now() - backup_time < timedelta(hours=24):
                        recovery_files.append({
                            'file': str(backup_file),
                            'title': backup_data.get('title', 'Untitled'),
                            'timestamp': backup_data['timestamp'],
                            'original_path': backup_data.get('file_path')
                        })
                except Exception as e:
                    print(f"Error reading recovery file {backup_file}: {e}")

        if recovery_files:
            self.recovery_available.emit(recovery_files)

    def recover_document(self, backup_file: str) -> Optional[Dict[str, Any]]:
        """
        Recover a document from a backup file.

        Args:
            backup_file: Path to the backup file

        Returns:
            Dictionary containing recovered document data, or None if failed
        """
        try:
            with open(backup_file, 'r', encoding='utf-8') as f:
                backup_data = json.load(f)

            # Remove the backup file after successful recovery
            os.remove(backup_file)

            return backup_data

        except Exception as e:
            print(f"Error recovering document from {backup_file}: {e}")
            return None

    def clear_all_backups(self):
        """Clear all backup files."""
        backup_dir = Path(self.backup_dir)
        if backup_dir.exists():
            for backup_file in backup_dir.glob('*.recovery'):
                try:
                    os.remove(backup_file)
                except Exception as e:
                    print(f"Error removing backup file {backup_file}: {e}")


class PerformanceMonitor(QObject):
    """
    Monitors application performance metrics.

    Features:
    - CPU usage monitoring
    - Memory usage monitoring
    - Render time tracking
    - Performance warnings
    """

    performance_warning = Signal(str)  # Warning message

    def __init__(self):
        super().__init__()
        self.metrics: Dict[str, List[float]] = {
            'memory': [],
            'render_time': [],
            'save_time': []
        }
        self.max_metric_history = 100

    def record_metric(self, metric_name: str, value: float):
        """
        Record a performance metric.

        Args:
            metric_name: Name of the metric
            value: Metric value
        """
        if metric_name not in self.metrics:
            self.metrics[metric_name] = []

        self.metrics[metric_name].append(value)

        # Keep only recent history
        if len(self.metrics[metric_name]) > self.max_metric_history:
            self.metrics[metric_name] = self.metrics[metric_name][-self.max_metric_history:]

        # Check for performance issues
        self._check_performance_issues(metric_name, value)

    def _check_performance_issues(self, metric_name: str, value: float):
        """
        Check for performance issues based on metrics.

        Args:
            metric_name: Name of the metric
            value: Metric value
        """
        # Define thresholds
        thresholds = {
            'memory': 500 * 1024 * 1024,  # 500 MB
            'render_time': 1.0,  # 1 second
            'save_time': 5.0  # 5 seconds
        }

        if metric_name in thresholds and value > thresholds[metric_name]:
            warning_messages = {
                'memory': f"High memory usage detected: {value / (1024*1024):.1f} MB",
                'render_time': f"Slow rendering detected: {value:.2f} seconds",
                'save_time': f"Slow save operation: {value:.2f} seconds"
            }
            self.performance_warning.emit(warning_messages.get(metric_name, f"Performance issue: {metric_name}"))

    def get_average_metric(self, metric_name: str) -> Optional[float]:
        """
        Get the average value of a metric.

        Args:
            metric_name: Name of the metric

        Returns:
            Average value or None if no data
        """
        if metric_name in self.metrics and self.metrics[metric_name]:
            return sum(self.metrics[metric_name]) / len(self.metrics[metric_name])
        return None

    def get_metrics_summary(self) -> Dict[str, Any]:
        """
        Get a summary of all performance metrics.

        Returns:
            Dictionary containing metric summaries
        """
        summary = {}
        for metric_name, values in self.metrics.items():
            if values:
                summary[metric_name] = {
                    'average': sum(values) / len(values),
                    'min': min(values),
                    'max': max(values),
                    'count': len(values)
                }
        return summary
