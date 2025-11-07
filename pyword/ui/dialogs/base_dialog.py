from PySide6.QtWidgets import QDialog, QDialogButtonBox, QVBoxLayout, QWidget
from PySide6.QtCore import Signal, Qt

class BaseDialog(QDialog):
    """Base dialog class with common functionality for all dialogs."""
    
    # Signal emitted when dialog is accepted
    accepted = Signal(dict)  # Dictionary of properties/values
    
    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the dialog UI. Should be implemented by subclasses."""
        self.layout = QVBoxLayout(self)
        
        # Content widget (to be added by subclasses)
        self.content = QWidget()
        self.layout.addWidget(self.content)
        
        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            Qt.Horizontal,
            self
        )
        self.button_box.accepted.connect(self.validate_and_accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box)
    
    def get_values(self):
        """Get the current values from the dialog's controls.
        
        Should be implemented by subclasses.
        Returns:
            dict: Dictionary of property-value pairs
        """
        return {}
    
    def set_values(self, values):
        """Set values to the dialog's controls.
        
        Args:
            values (dict): Dictionary of property-value pairs
        """
        pass
    
    def validate(self):
        """Validate the current input.
        
        Returns:
            tuple: (is_valid, error_message)
        """
        return True, ""
    
    def validate_and_accept(self):
        """Validate input and accept the dialog if valid."""
        is_valid, error = self.validate()
        if is_valid:
            self.accept()
        elif error:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Validation Error", error)
    
    def accept(self):
        """Accept the dialog and emit the accepted signal with the values."""
        values = self.get_values()
        self.accepted.emit(values)
        super().accept()
    
    def reject(self):
        """Reject the dialog."""
        super().reject()
    
    @classmethod
    def get_input(cls, title, initial_values=None, parent=None):
        """Static method to show the dialog and return the result.
        
        Args:
            title (str): Dialog title
            initial_values (dict, optional): Initial values for the dialog
            parent (QWidget, optional): Parent widget
            
        Returns:
            tuple: (accepted, values)
        """
        dialog = cls(title, parent)
        if initial_values:
            dialog.set_values(initial_values)
        
        result = dialog.exec_()
        if result == QDialog.Accepted:
            return True, dialog.get_values()
        return False, None
