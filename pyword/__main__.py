import sys
import os
from PySide6.QtWidgets import QApplication

# Handle both direct execution and module execution
if __name__ == "__main__" and __package__ is None:
    # Direct execution: python pyword/__main__.py
    # Add parent directory to path to allow absolute imports
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from pyword.core.application import WordProcessor
else:
    # Module execution: python -m pyword
    from .core.application import WordProcessor

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("PyWord")
    app.setOrganizationName("PyWord")
    
    # Set application style
    app.setStyle("Fusion")
    
    # Create and show main window
    window = WordProcessor()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
