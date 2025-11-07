import sys
from PySide6.QtWidgets import QApplication
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
